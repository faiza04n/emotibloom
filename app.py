import json
import logging
import os
import re
from datetime import datetime
from json import JSONDecodeError
from pathlib import Path

from flask import Flask, abort, jsonify, render_template, request, session
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"
STATIC_IMAGE_DIR = BASE_DIR / "static" / "images"
CATALOG_FILE = DATA_DIR / "curations.json"
ORDERS_FILE = DATA_DIR / "orders.xlsx"
_catalog_cache = {"mtime_ns": None, "data": None}

# Flask uses the curated catalog as the single source of truth; standalone
# flower data is intentionally not loaded anywhere in the application.
app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get(
    "SECRET_KEY",
    "emotibloom-dev-secret-key",
)


logging.basicConfig(level=logging.INFO)
app.logger.setLevel(logging.INFO)


class CatalogError(Exception):
    """Raised when the curated catalog is missing, malformed, or incomplete."""

    def __init__(self, message, status_code=500, details=None):
        super().__init__(message)
        self.status_code = status_code
        self.details = details or {}


def normalize_text(value):
    """Normalize text for tolerant matching without changing display labels."""
    return "".join(ch.lower() for ch in str(value) if ch.isalnum())


def load_json_file(path):
    """Load JSON with precise errors so catalog issues are easy to debug."""
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError as exc:
        raise CatalogError("Curated catalog file is missing", details={"path": str(path)}) from exc
    except JSONDecodeError as exc:
        raise CatalogError(
            "Curated catalog file contains invalid JSON",
            details={"path": str(path), "line": exc.lineno, "column": exc.colno},
        ) from exc


def validate_asset_path(image_path, required_folder, pointer, errors):
    """Validate that catalog image paths stay inside their expected folders."""
    if not isinstance(image_path, str) or not image_path.strip():
        errors.append(f"{pointer}: image_path is required")
        return

    relative = Path(image_path)
    if relative.is_absolute() or ".." in relative.parts:
        errors.append(f"{pointer}: image_path must be a safe relative path")
        return

    if not relative.parts or relative.parts[0] != required_folder:
        errors.append(f"{pointer}: image_path must start with {required_folder}/")
        return

    # Resolve paths before comparison to block traversal such as ../secrets.txt.
    root = (STATIC_IMAGE_DIR / required_folder).resolve()
    full_path = (STATIC_IMAGE_DIR / relative).resolve()
    try:
        full_path.relative_to(root)
    except ValueError:
        errors.append(f"{pointer}: image_path escapes the {required_folder} asset folder")
        return

    if not full_path.exists():
        errors.append(f"{pointer}: asset not found at static/images/{image_path}")


def validate_catalog(catalog):
    """Enforce the Emotion -> 3 bouquets -> 2 gifts hierarchy at startup."""
    errors = []
    emotions = catalog.get("emotions") if isinstance(catalog, dict) else None
    if not isinstance(emotions, list) or not emotions:
        raise CatalogError("Catalog must contain a non-empty emotions list")

    emotion_ids = set()
    for emotion_index, emotion in enumerate(emotions):
        pointer = f"emotions[{emotion_index}]"
        emotion_id = emotion.get("id")
        emotion_name = emotion.get("emotion")
        if not emotion_id or not emotion_name:
            errors.append(f"{pointer}: id and emotion are required")
        elif emotion_id in emotion_ids:
            errors.append(f"{pointer}: duplicate emotion id {emotion_id}")
        else:
            emotion_ids.add(emotion_id)

        bouquets = emotion.get("bouquets")
        if not isinstance(bouquets, list) or len(bouquets) != 3:
            errors.append(f"{pointer}: exactly 3 bouquet suggestions are required")
            continue

        bouquet_ids = set()
        for bouquet_index, bouquet in enumerate(bouquets):
            b_pointer = f"{pointer}.bouquets[{bouquet_index}]"
            for field in ("id", "bouquet_name", "image_path", "description", "price"):
                if field not in bouquet:
                    errors.append(f"{b_pointer}: {field} is required")
            if bouquet.get("id") in bouquet_ids:
                errors.append(f"{b_pointer}: duplicate bouquet id {bouquet.get('id')}")
            bouquet_ids.add(bouquet.get("id"))
            validate_asset_path(bouquet.get("image_path"), "bouquets", b_pointer, errors)

            gifts = bouquet.get("gifts")
            if not isinstance(gifts, list) or len(gifts) != 2:
                errors.append(f"{b_pointer}: exactly 2 gift options are required")
                continue

            gift_ids = set()
            for gift_index, gift in enumerate(gifts):
                g_pointer = f"{b_pointer}.gifts[{gift_index}]"
                for field in ("id", "gift_name", "image_path", "message", "price"):
                    if field not in gift:
                        errors.append(f"{g_pointer}: {field} is required")
                if gift.get("id") in gift_ids:
                    errors.append(f"{g_pointer}: duplicate gift id {gift.get('id')}")
                gift_ids.add(gift.get("id"))
                validate_asset_path(gift.get("image_path"), "gifts", g_pointer, errors)

                # Messages are deliberately two lines so cards remain readable.
                message_lines = [
                    line.strip()
                    for line in str(gift.get("message", "")).splitlines()
                    if line.strip()
                ]
                if len(message_lines) != 2:
                    errors.append(f"{g_pointer}: message must contain exactly 2 non-empty lines")

    if errors:
        raise CatalogError("Curated catalog validation failed", details={"errors": errors})


def load_catalog():
    """Load and validate the catalog, refreshing when curations.json changes."""
    try:
        mtime_ns = CATALOG_FILE.stat().st_mtime_ns
    except FileNotFoundError as exc:
        raise CatalogError(
            "Curated catalog file is missing",
            details={"path": str(CATALOG_FILE)},
        ) from exc

    if _catalog_cache["mtime_ns"] == mtime_ns and _catalog_cache["data"] is not None:
        return _catalog_cache["data"]

    catalog = load_json_file(CATALOG_FILE)
    validate_catalog(catalog)
    _catalog_cache["mtime_ns"] = mtime_ns
    _catalog_cache["data"] = catalog
    return catalog


def get_catalog():
    """Return the cached catalog while logging validation failures centrally."""
    try:
        return load_catalog()
    except CatalogError:
        app.logger.exception("Catalog could not be loaded")
        raise


def package_id(bouquet_id, gift_id):
    """Build the cart-safe package id used for bouquet + gift combinations."""
    return f"{bouquet_id}__{gift_id}"


def image_url(image_path):
    """Convert catalog-relative image paths into browser-ready static URLs."""
    return f"/static/images/{image_path}"


def serialize_gift(gift, bouquet):
    """Expose gift data with compatibility aliases expected by templates/JS."""
    return {
        "id": gift["id"],
        "gift_name": gift["gift_name"],
        "name": gift["gift_name"],
        "image_path": gift["image_path"],
        "image": gift["image_path"],
        "image_url": image_url(gift["image_path"]),
        "message": gift["message"],
        "price": float(gift["price"]),
        "package_id": package_id(bouquet["id"], gift["id"]),
        "package_price": round(float(bouquet["price"]) + float(gift["price"]), 2),
    }


def serialize_bouquet(bouquet):
    """Expose one bouquet and its two gift options in API/template shape."""
    gifts = [serialize_gift(gift, bouquet) for gift in bouquet["gifts"]]
    return {
        "id": bouquet["id"],
        "bouquet_name": bouquet["bouquet_name"],
        "name": bouquet["bouquet_name"],
        "image_path": bouquet["image_path"],
        "image": bouquet["image_path"],
        "image_url": image_url(bouquet["image_path"]),
        "description": bouquet["description"],
        "price": float(bouquet["price"]),
        "gifts": gifts,
    }


def serialize_emotion(emotion):
    """Expose a complete emotion curation with its three bouquets."""
    bouquets = [serialize_bouquet(bouquet) for bouquet in emotion["bouquets"]]
    unique_gifts = []
    seen_gifts = set()
    for bouquet in bouquets:
        for gift in bouquet["gifts"]:
            key = normalize_text(gift["gift_name"])
            if key in seen_gifts:
                continue
            seen_gifts.add(key)
            unique_gifts.append({
                **gift,
                "bouquet_id": bouquet["id"],
                "bouquet_name": bouquet["bouquet_name"],
            })

    return {
        "id": emotion["id"],
        "emotion": emotion["emotion"],
        "name": emotion["emotion"],
        "aliases": emotion.get("aliases", []),
        "color": emotion.get("color", "#ff3d7f"),
        "description": emotion.get("description", ""),
        "bouquets": bouquets,
        "unique_gifts": unique_gifts,
    }


def emotion_summaries():
    """Return serialized emotions for navigation, pages, and API responses."""
    return [serialize_emotion(emotion) for emotion in get_catalog()["emotions"]]


def find_emotion(query):
    """Match exact, prefix, then contains so partial searches like 'lov' work."""
    q_norm = normalize_text(query)
    if not q_norm:
        return None

    emotions = get_catalog()["emotions"]

    # First prefer exact matches against ids, display names, and aliases.
    for emotion in emotions:
        labels = [emotion["id"], emotion["emotion"], *emotion.get("aliases", [])]
        if q_norm in {normalize_text(label) for label in labels}:
            return emotion

    # Then support partial typing from the beginning of each label.
    for emotion in emotions:
        labels = [emotion["id"], emotion["emotion"], *emotion.get("aliases", [])]
        if any(normalize_text(label).startswith(q_norm) for label in labels):
            return emotion

    # Finally allow broader contains matches for phrases such as "thank".
    for emotion in emotions:
        labels = [emotion["id"], emotion["emotion"], *emotion.get("aliases", [])]
        if any(q_norm in normalize_text(label) for label in labels):
            return emotion
    return None


def flatten_bouquets():
    """Flatten emotion bouquets for grid pages and search without duplicating data."""
    items = []
    for emotion in get_catalog()["emotions"]:
        serialized_emotion = serialize_emotion(emotion)
        for bouquet in serialized_emotion["bouquets"]:
            items.append({
                **bouquet,
                "emotion": serialized_emotion["emotion"],
                "emotions": [serialized_emotion["emotion"]],
            })
    return items


def flatten_gifts():
    """Flatten gifts once per emotion so repeated bouquet pairings do not duplicate cards."""
    items = []
    seen = set()
    for emotion in get_catalog()["emotions"]:
        serialized_emotion = serialize_emotion(emotion)
        for bouquet in serialized_emotion["bouquets"]:
            for gift in bouquet["gifts"]:
                key = (serialized_emotion["id"], normalize_text(gift["gift_name"]))
                if key in seen:
                    continue
                seen.add(key)
                items.append({
                    **gift,
                    "emotion": serialized_emotion["emotion"],
                    "emotions": [serialized_emotion["emotion"]],
                    "bouquet_id": bouquet["id"],
                    "bouquet_name": bouquet["bouquet_name"],
                })
    return items


def flatten_packages():
    """Flatten every bouquet + gift pairing into add-to-cart package entries."""
    packages = []
    for bouquet in flatten_bouquets():
        for gift in bouquet["gifts"]:
            packages.append({
                "id": gift["package_id"],
                "name": f"{bouquet['bouquet_name']} + {gift['gift_name']}",
                "emotion": bouquet["emotion"],
                "emotions": [bouquet["emotion"]],
                "bouquet_name": bouquet["bouquet_name"],
                "gift_name": gift["gift_name"],
                "message": gift["message"],
                "image": bouquet["image"],
                "image_path": bouquet["image_path"],
                "price": gift["package_price"],
            })
    return packages


def find_catalog_item(category, item_id):
    """Look up a cart item from the appropriate flattened catalog view."""
    sources = {
        "bouquet": flatten_bouquets,
        "gift": flatten_gifts,
        "package": flatten_packages,
    }
    source = sources.get(category)
    if not source:
        return None
    return next((item for item in source() if str(item["id"]) == str(item_id)), None)


def normalize_category(category):
    """Accept singular/plural category labels from buttons and API clients."""
    aliases = {
        "bouquet": "bouquet",
        "bouquets": "bouquet",
        "gift": "gift",
        "gifts": "gift",
        "package": "package",
        "packages": "package",
        "custom": "custom",
    }
    return aliases.get(str(category).lower())


def cart_total(cart):
    """Calculate a rounded cart total from session-backed cart rows."""
    return round(sum(item["price"] * item["qty"] for item in cart), 2)


def get_cart():
    """Return the current session cart, creating it on first use."""
    if "cart" not in session:
        session["cart"] = []
    return session["cart"]


def save_cart(cart):
    """Persist cart mutations back into the Flask session."""
    session["cart"] = cart
    session.modified = True


def create_orders_file():
    """Create orders.xlsx if it doesn't exist with proper headers."""
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    if ORDERS_FILE.exists():
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "Orders"

    headers = [
        "Order ID", "Date", "First Name", "Last Name", "Email", "Phone",
        "Address", "City", "Postal Code", "Products", "Quantity", "Total", "Status"
    ]
    ws.append(headers)

    header_fill = PatternFill(start_color="FF3D7F", end_color="FF3D7F", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 12
    ws.column_dimensions["E"].width = 20
    ws.column_dimensions["F"].width = 14
    ws.column_dimensions["G"].width = 20
    ws.column_dimensions["H"].width = 12
    ws.column_dimensions["I"].width = 12
    ws.column_dimensions["J"].width = 30
    ws.column_dimensions["K"].width = 10
    ws.column_dimensions["L"].width = 10
    ws.column_dimensions["M"].width = 12

    try:
        wb.save(ORDERS_FILE)
    finally:
        wb.close()
    app.logger.info("Created orders.xlsx")


def generate_order_id():
    """Generate next order ID in format EMB-XXXX, starting from EMB-1001."""
    create_orders_file()

    try:
        wb = load_workbook(ORDERS_FILE)
        try:
            ws = wb.active
            next_id = 1001
            for row in range(ws.max_row, 1, -1):
                order_id = ws.cell(row=row, column=1).value
                if isinstance(order_id, str) and order_id.startswith("EMB-"):
                    try:
                        next_id = int(order_id.split("-", 1)[1]) + 1
                        break
                    except ValueError:
                        continue
        finally:
            wb.close()

        return f"EMB-{next_id}"
    except Exception as e:
        app.logger.error("Error generating order ID: %s", e)
        return f"EMB-{datetime.now().strftime('%m%d%H%M%S')}"


def validate_customer_data(data):
    """Validate customer information form data."""
    errors = []

    required_fields = [
        "first_name",
        "last_name",
        "email",
        "phone",
        "address",
        "city",
        "postal_code",
    ]
    for field in required_fields:
        value = (data.get(field) or "").strip()
        if not value:
            errors.append(f"{field.replace('_', ' ').title()} is required")

    email = (data.get("email") or "").strip()
    if email and not re.match(r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$", email):
        errors.append("Invalid email address")

    phone = (data.get("phone") or "").strip()
    if phone and not re.match(r"^[\d\s\-+()]+$", phone):
        errors.append("Invalid phone number")

    return errors


def save_order(customer_data, cart_items, total):
    """Save order to Excel file."""
    create_orders_file()

    order_id = generate_order_id()
    date_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    product_list = ", ".join([
        f"{item['name']} x{item['qty']}" for item in cart_items
    ])
    quantity_total = sum(item["qty"] for item in cart_items)

    try:
        wb = load_workbook(ORDERS_FILE)
        try:
            ws = wb.active

            order_row = [
                order_id,
                date_str,
                customer_data.get("first_name", "").strip(),
                customer_data.get("last_name", "").strip(),
                customer_data.get("email", "").strip(),
                customer_data.get("phone", "").strip(),
                customer_data.get("address", "").strip(),
                customer_data.get("city", "").strip(),
                customer_data.get("postal_code", "").strip(),
                product_list,
                quantity_total,
                round(float(total), 2),
                "Pending",
            ]

            ws.append(order_row)
            wb.save(ORDERS_FILE)
        finally:
            wb.close()
        app.logger.info("Order saved: %s", order_id)
        return order_id
    except Exception as e:
        app.logger.error("Error saving order: %s", e)
        raise


def get_orders():
    """Retrieve all orders from Excel file, sorted by date (newest first)."""
    create_orders_file()

    orders = []
    try:
        wb = load_workbook(ORDERS_FILE)
        try:
            ws = wb.active

            for row in ws.iter_rows(min_row=2, values_only=True):
                padded = list(row or []) + [None] * 13
                if padded[0]:
                    order = {
                        "order_id": padded[0],
                        "date": padded[1],
                        "first_name": padded[2],
                        "last_name": padded[3],
                        "email": padded[4],
                        "phone": padded[5],
                        "address": padded[6],
                        "city": padded[7],
                        "postal_code": padded[8],
                        "products": padded[9],
                        "quantity": padded[10],
                        "total": padded[11],
                        "status": padded[12] or "Pending",
                    }
                    orders.append(order)
        finally:
            wb.close()

        orders.reverse()
        return orders
    except Exception as e:
        app.logger.error("Error retrieving orders: %s", e)
        return []


def api_error(message, status_code=400, details=None):
    """Return consistent API errors for frontend fetch handlers."""
    payload = {"ok": False, "error": message}
    if details:
        payload["details"] = details
    return jsonify(payload), status_code


def cart_payload(cart):
    """Return consistent numeric cart metadata for API responses."""
    return {
        "cart_count": int(sum(int(i.get("qty", 0)) for i in cart)),
        "cart_total": float(cart_total(cart)),
    }


@app.context_processor
def inject_globals():
    """Expose cart and emotion navigation data to every template."""
    cart = session.get("cart", [])
    try:
        nav_emotions = emotion_summaries()
    except CatalogError:
        nav_emotions = []
    return {
        "cart_count": cart_payload(cart)["cart_count"],
        "nav_emotions": nav_emotions,
    }


@app.route("/")
def home():
    """Render the homepage with a small featured preview from each module."""
    emotions = emotion_summaries()
    return render_template(
        "home.html",
        emotions=emotions,
        featured_bouquets=flatten_bouquets()[:3],
        featured_gifts=flatten_gifts()[:3],
        featured_packages=flatten_packages()[:3],
        wishcards=flatten_gifts()[:3],
    )


@app.route("/emotions")
def emotions_page():
    """Render all emotions or a selected emotion's full curation flow."""
    emotions = emotion_summaries()
    selected_key = request.args.get("emotion")
    selected_curation = None
    if selected_key:
        selected = find_emotion(selected_key)
        if not selected:
            app.logger.info("Invalid emotion page selection: %s", selected_key)
            abort(404)
        selected_curation = serialize_emotion(selected)
    return render_template(
        "emotions.html",
        emotions=emotions,
        selected=selected_curation["emotion"] if selected_curation else None,
        curation=selected_curation,
    )


@app.route("/bouquets")
def bouquets_page():
    """Render every curated bouquet from the unified catalog."""
    return render_template("bouquets.html", bouquets=flatten_bouquets())


@app.route("/gifts")
def gifts_page():
    """Render every gift/message option from the unified catalog."""
    return render_template("gifts.html", gifts=flatten_gifts())


@app.route("/customize", methods=["GET", "POST"])
def customize_page():
    """Render the client-side atelier builder using the full emotion hierarchy."""
    return render_template("customize.html", curations=emotion_summaries())


@app.route("/about")
def about_page():
    return render_template("about.html")


@app.route("/cart")
def cart_page():
    """Render the session-backed demo cart."""
    cart = get_cart()
    return render_template("cart.html", cart=cart, total=cart_total(cart))


@app.route("/api/order/place", methods=["POST"])
def api_place_order():
    """Place an order with customer information and cart items."""
    payload = request.get_json(silent=True) or {}
    cart = get_cart()

    if not cart:
        return api_error("Cart is empty", 400)

    customer_data = {
        "first_name": payload.get("first_name", "").strip(),
        "last_name": payload.get("last_name", "").strip(),
        "email": payload.get("email", "").strip(),
        "phone": payload.get("phone", "").strip(),
        "address": payload.get("address", "").strip(),
        "city": payload.get("city", "").strip(),
        "postal_code": payload.get("postal_code", "").strip(),
    }

    validation_errors = validate_customer_data(customer_data)
    if validation_errors:
        return api_error("Validation failed", 400, {"errors": validation_errors})

    try:
        total = cart_total(cart)
        order_id = save_order(customer_data, cart, total)
        order_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        products = [
            {
                "name": item.get("name", "Cart item"),
                "category": item.get("category", "item"),
                "qty": int(item.get("qty", 1)),
                "price": float(item.get("price", 0)),
                "line_total": round(float(item.get("price", 0)) * int(item.get("qty", 1)), 2),
            }
            for item in cart
        ]

        # Build wish_cards array from gifts in the cart.
        wish_cards = []
        for item in cart:
            if item["category"] == "gift":
                found = find_catalog_item("gift", item["id"].removeprefix("gift-"))
                if found:
                    for _ in range(item.get("qty", 1)):
                        wish_cards.append({
                            "gift_name": found["name"],
                            "message": found["message"],
                        })
            elif item["category"] == "package":
                found = find_catalog_item("package", item["id"].removeprefix("package-"))
                if found:
                    for _ in range(item.get("qty", 1)):
                        wish_cards.append({
                            "gift_name": found["gift_name"],
                            "message": found["message"],
                        })

        session["cart"] = []
        session.modified = True

        return jsonify({
            "ok": True,
            "order_id": order_id,
            "date": order_date,
            "status": "Pending",
            "products": products,
            "total": float(total),
            "cart_count": 0,
            "cart_total": 0.0,
            "message": "Order placed successfully",
            "wish_cards": wish_cards,
        })
    except Exception as e:
        app.logger.error("Error placing order: %s", e)
        return api_error("Failed to place order", 500)


@app.route("/admin/orders")
def admin_orders_page():
    """Display all orders in admin panel."""
    admin_key = os.environ.get("EMOTIBLOOM_ADMIN_KEY")
    supplied_key = request.headers.get("X-Admin-Key") or request.args.get("key")
    if admin_key and supplied_key != admin_key:
        abort(404)

    orders = get_orders()
    stats = {
        "total_orders": len(orders),
        "pending_orders": sum(1 for o in orders if o.get("status") == "Pending"),
        "processing_orders": sum(1 for o in orders if o.get("status") == "Processing"),
        "delivered_orders": sum(1 for o in orders if o.get("status") == "Delivered"),
    }
    return render_template("admin-orders.html", orders=orders, stats=stats)


@app.route("/api/catalog/validate")
def api_catalog_validate():
    """Health-check endpoint for validating catalog counts and schema version."""
    catalog = get_catalog()
    bouquet_count = sum(len(emotion["bouquets"]) for emotion in catalog["emotions"])
    gift_count = sum(
        len(bouquet["gifts"])
        for emotion in catalog["emotions"]
        for bouquet in emotion["bouquets"]
    )
    return jsonify({
        "ok": True,
        "schema_version": catalog.get("schema_version"),
        "emotion_count": len(catalog["emotions"]),
        "bouquet_count": bouquet_count,
        "gift_count": len(flatten_gifts()),
        "gift_option_count": gift_count,
    })


@app.route("/api/emotions")
def api_emotions():
    """Serve the full emotion hierarchy for frontend integrations."""
    return jsonify({"ok": True, "emotions": emotion_summaries()})


@app.route("/api/emotions/<emotion_key>")
def api_emotion_detail(emotion_key):
    """Serve one emotion curation by id, name, alias, or partial match."""
    emotion = find_emotion(emotion_key)
    if not emotion:
        return api_error("Invalid emotion selection", 404, {"emotion": emotion_key})
    return jsonify({"ok": True, "emotion": serialize_emotion(emotion)})


@app.route("/api/search")
def api_search():
    """Serve mixed search results while leaving recommendations to their own API."""
    q = request.args.get("q", "").strip()
    q_norm = normalize_text(q)
    results = {"emotions": [], "bouquets": [], "gifts": []}
    if not q_norm:
        return jsonify(results)

    for emotion in emotion_summaries():
        haystack = " ".join(
            [
                emotion["emotion"],
                emotion["description"],
                *emotion.get("aliases", []),
            ]
        )
        if q_norm in normalize_text(haystack):
            results["emotions"].append({
                "id": emotion["id"],
                "name": emotion["emotion"],
                "description": emotion["description"],
            })

    for bouquet in flatten_bouquets():
        haystack = " ".join([bouquet["name"], bouquet["description"], bouquet["emotion"]])
        if q_norm in normalize_text(haystack):
            results["bouquets"].append(bouquet)

    for gift in flatten_gifts():
        haystack = " ".join([gift["name"], gift["message"], gift["emotion"], gift["bouquet_name"]])
        if q_norm in normalize_text(haystack):
            results["gifts"].append(gift)

    return jsonify(results)


@app.route("/api/emotion-recommendation")
def api_emotion_recommendation():
    """Return one complete curation when a search query maps to an emotion."""
    q = request.args.get("q", "").strip()
    if not q:
        return jsonify({"matched_emotion": None, "recommendation": None})

    emotion = find_emotion(q)
    if not emotion:
        return jsonify({"matched_emotion": None, "recommendation": None})

    serialized = serialize_emotion(emotion)
    return jsonify({
        "matched_emotion": serialized["emotion"],
        "recommendation": serialized,
    })


@app.route("/api/cart/add", methods=["POST"])
def api_cart_add():
    """Add catalog or custom builder items to the session cart."""
    payload = request.get_json(silent=True) or {}
    category = normalize_category(payload.get("category"))
    item_id = payload.get("id")
    try:
        qty = int(payload.get("qty", 1))
    except (TypeError, ValueError):
        return api_error("Quantity must be a number", 400)
    if qty < 1 or qty > 99:
        return api_error("Quantity must be between 1 and 99", 400)

    if category not in ("bouquet", "gift", "package", "custom"):
        return api_error("Invalid category", 400)

    if category == "custom":
        # Custom items are accepted only with a positive server-validated price.
        try:
            price = float(payload.get("price", 0))
        except (TypeError, ValueError):
            return api_error("Custom item price must be a number", 400)
        if price <= 0:
            return api_error("Custom item price must be greater than zero", 400)
        item = {
            "id": f"custom-{len(get_cart()) + 1}",
            "name": payload.get("name", "Custom EmotiBloom Atelier"),
            "price": round(price, 2),
            "category": "custom",
            "details": payload.get("details", {}),
            "qty": qty,
        }
    else:
        found = find_catalog_item(category, item_id)
        if not found:
            return api_error("Item not found", 404, {"category": category, "id": item_id})
        item = {
            "id": f"{category}-{found['id']}",
            "name": found["name"],
            "price": float(found["price"]),
            "category": category,
            "qty": qty,
        }

    cart = get_cart()
    for existing in cart:
        # Merge duplicate cart rows instead of creating repeated entries.
        if existing["id"] == item["id"]:
            if existing["qty"] + qty > 99:
                return api_error("Quantity must be between 1 and 99", 400)
            existing["qty"] += qty
            break
    else:
        cart.append(item)
    save_cart(cart)
    return jsonify({"ok": True, **cart_payload(cart)})


@app.route("/api/cart/update", methods=["POST"])
def api_cart_update():
    """Update an existing cart row quantity within safe demo-cart bounds."""
    payload = request.get_json(silent=True) or {}
    item_id = payload.get("id")
    try:
        qty = int(payload.get("qty", 1))
    except (TypeError, ValueError):
        return api_error("Quantity must be a number", 400)
    if qty < 1 or qty > 99:
        return api_error("Quantity must be between 1 and 99", 400)

    cart = get_cart()
    matched = False
    for item in cart:
        if item["id"] == item_id:
            item["qty"] = qty
            matched = True
            break
    if not matched:
        return api_error("Cart item not found", 404, {"id": item_id})
    save_cart(cart)
    return jsonify({"ok": True, "total": cart_total(cart), **cart_payload(cart)})


@app.route("/api/cart/remove", methods=["POST"])
def api_cart_remove():
    """Remove a cart row by id and return fresh totals for the UI."""
    payload = request.get_json(silent=True) or {}
    item_id = payload.get("id")
    cart = [i for i in get_cart() if i["id"] != item_id]
    save_cart(cart)
    return jsonify({
        "ok": True,
        "total": cart_total(cart),
        **cart_payload(cart),
    })


@app.errorhandler(CatalogError)
def handle_catalog_error(error):
    """Render API JSON or a friendly page when catalog validation fails."""
    app.logger.error("Catalog error: %s", error)
    if request.path.startswith("/api/"):
        return api_error(str(error), error.status_code, error.details)
    return render_template("404.html"), error.status_code


@app.errorhandler(404)
def not_found(_):
    """Render the shared not-found page for invalid routes and selections."""
    if request.path.startswith("/api/"):
        return api_error("Not found", 404)
    return render_template("404.html"), 404


@app.errorhandler(500)
def server_error(error):
    """Log unexpected failures while keeping API/page responses consistent."""
    app.logger.exception("Unhandled server error: %s", error)
    if request.path.startswith("/api/"):
        return api_error("Internal server error", 500)
    return render_template("404.html"), 500


get_catalog()


if __name__ == "__main__":
    debug_enabled = os.environ.get("FLASK_DEBUG", "").lower() in {"1", "true", "yes"}
    app.run(debug=debug_enabled)
