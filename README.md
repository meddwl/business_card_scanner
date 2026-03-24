# 📇 Business Card Scanner

> An AI-powered Python script that extracts contact information from business card photos and saves them to a local Excel file automatically.

---

## ✨ Features

- 📸 Reads business card images from a local `input/` folder
- 🤖 Uses **Gemini 2.5 Flash** Vision to extract: Full Name, Position, Company, Phone, Email
- 💾 Appends each contact as a new row in a local **Excel file** (`contacts.xlsx`)
- 📁 Moves processed images to a `processed/` folder
- 🔁 Supports batch processing (multiple cards at once)
- 🛡️ Handles missing fields by returning an empty string

---

## 🗂️ Project Structure

```
business-card-scanner/
├── input/                      ← Drop your business card images here
├── processed/                  ← Processed images land here automatically
├── contacts.xlsx               ← Created automatically on first run
├── business_card_scanner.py    ← Main script
├── .env                        ← Your API key (never commit!)
├── .env.example                ← Template — safe to commit
├── requirements.txt
└── README.md
```

---

## ⚙️ Setup

### 1. Clone & install dependencies

```bash
git clone https://github.com/meddwl/business-card-scanner.git
cd business-card-scanner
pip install -r requirements.txt
```

### 2. Set up your `.env` file

```bash
cp .env.example .env
# Then edit .env and fill in your Gemini API key
```

### 3. Gemini API key

1. Go to [aistudio.google.com/app/apikey](https://aistudio.google.com/app/apikey)
2. Create an API key (free to start)
3. Paste it into `.env` as `GEMINI_API_KEY`

That's it — no service accounts, no cloud setup needed.

Free Tier Limits: 5 business cards per minute and 20 per day on the Gemini 2.5 Flash free tier.

---

## 🚀 Usage

1. Drop one or more business card images (`.jpg`, `.jpeg`, `.png`, `.webp`, `.gif`, `.heic`, `.heif`) into the `input/` folder
2. Run the script:

```bash
python business_card_scanner.py
```

3. Open `contacts.xlsx` to see your extracted contacts!

---

## 📊 Output format (Excel)

| Full Name | Position | Company   | Phone       | Email       | Processed At        | Source File  |
|-----------|----------|-----------|-------------|-------------|---------------------|--------------|
| Alina D   | CTO      | Work Corp | +4123456789 | ad@work.com | 2026-03-24 10:22:01 | card_001.jpg |

The file is created automatically on the first run. Each new run appends rows — it never overwrites existing data.

---

## 🛠️ Tech Stack

| Tool | Purpose |
|------|---------|
| [google-genai](https://googleapis.github.io/js-genai/) | Gemini 2.5 Flash Vision API |
| [openpyxl](https://openpyxl.readthedocs.io) | Read/write Excel `.xlsx` files |
| [python-dotenv](https://github.com/theskumar/python-dotenv) | Environment variable management |

---

## 🔒 Security

- **Never commit** `.env` to version control — it contains your API key
- `contacts.xlsx` and image folders are listed in `.gitignore`

---

## 📄 License

MIT
