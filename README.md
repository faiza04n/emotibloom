# EmotiBloom

EmotiBloom is an emotion-based flower and gift recommendation system built with Flask. Users choose or search an emotion, then receive curated bouquet, gift, and two-line message pairings designed for thoughtful gifting.

## Features

- Emotion-led recommendation flow with 10 curated emotions
- Three bouquet suggestions per emotion
- Two gift and message options per bouquet
- Search API for emotions, bouquets, and gifts
- Session-based cart with add, update, and remove actions
- Custom atelier builder route
- Order placement with Excel-backed order storage
- Admin orders page with optional `EMOTIBLOOM_ADMIN_KEY` protection
- Responsive Flask templates with organized static assets

## Technologies Used

- Python
- Flask
- Jinja2
- JavaScript
- HTML5
- CSS3
- JSON
- openpyxl
- Pillow

## Project Structure

```text
/
├── app.py
├── requirements.txt
├── README.md
├── LICENSE
├── .gitignore
├── data/
│   └── curations.json
├── static/
│   ├── css/
│   ├── js/
│   └── images/
│       ├── backgrounds/
│       ├── bouquets/
│       ├── gifts/
│       └── icons/
├── templates/
├── screenshots/
└── tools/
```

## Installation

1. Clone the repository:

```bash
git clone https://github.com/your-username/emotibloom.git
cd emotibloom
```

2. Create and activate a virtual environment:

```bash
python -m venv venv
venv\Scripts\activate
```

On macOS/Linux:

```bash
source venv/bin/activate
```

3. Install dependencies:

```bash
pip install -r requirements.txt
```

4. Configure optional environment variables:

```bash
set SECRET_KEY=change-this-in-production
set EMOTIBLOOM_ADMIN_KEY=your-admin-key
```

On macOS/Linux:

```bash
export SECRET_KEY=change-this-in-production
export EMOTIBLOOM_ADMIN_KEY=your-admin-key
```

## Usage

Run the Flask application:

```bash
python app.py
```

Open the app in your browser:

```text
http://127.0.0.1:5000
```

Useful routes:

- `/` - homepage and emotion search
- `/emotions` - full emotion catalog
- `/bouquets` - bouquet catalog
- `/gifts` - gift catalog
- `/customize` - custom atelier builder
- `/cart` - shopping cart
- `/admin/orders` - order dashboard
- `/api/catalog/validate` - catalog health check

## Screenshots

Screenshots should be added to the `screenshots/` folder before publishing.

Suggested images:

- `screenshots/home-page.png`
- `screenshots/emotion-curation.png`
- `screenshots/bouquets-page.png`
- `screenshots/gift-card-message.png`
- `screenshots/cart-checkout.png`
- `screenshots/admin-orders.png`

## Data Model

The recommendation catalog lives in `data/curations.json`.

```text
Emotion
  -> 3 curated bouquets
    -> 2 gift options
      -> 2-line heartfelt message
```

Flask validates the catalog and required assets at startup so missing images, malformed JSON, and incomplete emotion data are caught early.

## Future Improvements

- Add user authentication for the admin dashboard
- Replace Excel order storage with a production database
- Add automated tests for API and cart flows
- Add payment integration
- Add deployment configuration for a cloud host

## License

This project is licensed under the MIT License. See `LICENSE` for details.

## Author

Muhammad Faizan
