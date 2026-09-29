# Text-Preprocessing-Web-Wikipedia-Provinsi-di-Indonesia
import requests
from bs4 import BeautifulSoup
import pandas as pd
import re
import string
import nltk
import csv
import matplotlib.pyplot as plt
from wordcloud import WordCloud

# ==========================================
# DOWNLOAD STOPWORDS NLTK
# ==========================================
nltk.download('stopwords', quiet=True)

from nltk.corpus import stopwords
stop_words = set(stopwords.words('english'))

# Tambahan stopword Bahasa Indonesia
stop_words.update([
    'dan', 'di', 'ke', 'dari', 'yang', 'untuk', 'dengan',
    'pada', 'adalah', 'ini', 'itu', 'atau', 'sebagai',
    'provinsi', 'province'
])

# ==========================================
# LANGKAH 1: SCRAPING TABEL WIKIPEDIA
# ==========================================

url = "https://en.wikipedia.org/wiki/Provinces_of_Indonesia"

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'
}

response = requests.get(url, headers=headers)

soup = BeautifulSoup(response.content, 'html.parser')

tables = soup.find_all('table', class_='wikitable')

target_table = None

for table in tables:
    if "Banda Aceh" in table.text or "Aceh" in table.text:
        target_table = table
        break

if target_table is None:
    print("Tabel provinsi tidak ditemukan.")
    exit()

data = []

rows = target_table.find_all('tr')

for row in rows:

    cols = row.find_all('td')

    if not cols:
        continue

    # Membersihkan isi setiap cell
    row_text = [
        re.sub(r'\[.*?\]', '', c.text).strip()
        for c in cols
    ]

    # Hanya mengambil baris yang memiliki kode angka
    if len(row_text) >= 7 and row_text[0].isdigit():

        iso_code = row_text[0]
        abbr = row_text[1]

        # Mencari bagian numerik
        numeric_parts = [
            t for t in row_text
            if re.match(r'^\d[\d,\s\.]*$', t)
        ]

        # Mencari bagian teks
        text_parts = [
            t for t in row_text
            if not re.match(r'^\d[\d,\s\.]*$', t)
        ]

        province_eng = text_parts[1] if len(text_parts) > 1 else ""
        province_ind = text_parts[2] if len(text_parts) > 2 else ""
        capital = text_parts[3] if len(text_parts) > 3 else ""

        # Area dan Population
        nums = [
            n for n in numeric_parts
            if n != iso_code
        ]

        area = nums[0] if len(nums) > 0 else ""
        population = nums[1] if len(nums) > 1 else ""

        data.append([
            iso_code,
            abbr,
            province_eng,
            province_ind,
            capital,
            area,
            population
        ])

# ==========================================
# MEMBUAT DATAFRAME
# ==========================================

df = pd.DataFrame(
    data,
    columns=[
        'Code',
        'Abbreviation',
        'Province',
        'Indonesian_Name',
        'Capital',
        'Area_km2',
        'Population'
    ]
)

# Nomor urut 1-38
df.insert(0, 'No', range(1, len(df) + 1))

# ==========================================
# LANGKAH 2: TEXT PROCESSING
# ==========================================

def nlp_text_processing(text):

    if pd.isna(text) or text == "":
        return ""

    text = str(text)

    # 1. Remove HTML Tags
    text = re.sub(r'<.*?>', '', text)

    # 2. Remove URLs
    text = re.sub(
        r'https?://\S+|www\.\S+',
        '',
        text
    )

    # 3. Remove Footnotes
    text = re.sub(
        r'\[.*?\]',
        '',
        text
    )

    # 4. Remove Emoji
    text = text.encode(
        'ascii',
        'ignore'
    ).decode('utf-8')

    # 5. Lowercase
    text = text.lower()

    # 6. Remove punctuation
    text = text.translate(
        str.maketrans(
            '',
            '',
            string.punctuation
        )
    )

    # 7. Stopword Removal
    words = text.split()

    filtered_words = [
        w for w in words
        if w not in stop_words
    ]

    return " ".join(filtered_words)


def process_numeric_text(val):

    if pd.isna(val) or val == "":
        return None

    clean_val = re.sub(
        r'[^\d]',
        '',
        str(val)
    )

    return int(clean_val) if clean_val != '' else None


# ==========================================
# MENERAPKAN TEXT PROCESSING
# ==========================================

for col in df.columns:

    if col in ['Area_km2', 'Population']:

        df[col] = df[col].apply(
            process_numeric_text
        )

    elif col not in [
        'No',
        'Code',
        'Abbreviation'
    ]:

        df[col] = df[col].apply(
            nlp_text_processing
        )


print("\n--- DATA SETELAH TEXT PROCESSING ---")

display(df.head(10))


# ==========================================
# LANGKAH 3: MEMBUAT WORD CLOUD
# ==========================================

# Gabungkan seluruh teks dari kolom teks
text_columns = [
    'Province',
    'Indonesian_Name',
    'Capital'
]

all_text = " ".join(
    df[text_columns]
    .fillna("")
    .astype(str)
    .values.flatten()
)

# Membuat Word Cloud
wordcloud = WordCloud(
    width=1000,
    height=500,
    background_color='white',
    max_words=100,
    min_font_size=10,
    collocations=False
).generate(all_text)

# ==========================================
# MENAMPILKAN WORD CLOUD
# ==========================================

plt.figure(
    figsize=(14, 7)
)

plt.imshow(
    wordcloud,
    interpolation='bilinear'
)

plt.axis('off')

plt.title(
    'Word Cloud Data Provinsi Indonesia',
    fontsize=20,
    pad=20
)

plt.tight_layout()

# Simpan gambar
wordcloud_file = "WordCloud_Provinsi_Indonesia.png"

plt.savefig(
    wordcloud_file,
    dpi=300,
    bbox_inches='tight'
)

plt.show()


# ==========================================
# LANGKAH 4: EXPORT CSV & EXCEL
# ==========================================

excel_file = "Data_Provinsi_Cleaned.xlsx"
csv_file = "Data_Provinsi_Cleaned.csv"


# ------------------------------------------
# SIMPAN CSV
# ------------------------------------------

df.to_csv(
    csv_file,
    index=False,
    encoding='utf-8-sig',
    quoting=csv.QUOTE_MINIMAL
)


# ------------------------------------------
# SIMPAN EXCEL
# ------------------------------------------

with pd.ExcelWriter(
    excel_file,
    engine='openpyxl'
) as writer:

    df.to_excel(
        writer,
        index=False,
        sheet_name='Provinces'
    )

    workbook = writer.book
    worksheet = writer.sheets['Provinces']

    from openpyxl.styles import (
        Font,
        PatternFill,
        Alignment,
        Border,
        Side
    )

    # Header
    header_fill = PatternFill(
        start_color="1F4E79",
        end_color="1F4E79",
        fill_type="solid"
    )

    header_font = Font(
        name="Calibri",
        size=11,
        bold=True,
        color="FFFFFF"
    )

    thin_border = Border(
        left=Side(
            style='thin',
            color='D9D9D9'
        ),
        right=Side(
            style='thin',
            color='D9D9D9'
        ),
        top=Side(
            style='thin',
            color='D9D9D9'
        ),
        bottom=Side(
            style='thin',
            color='D9D9D9'
        )
    )

    # Format Header
    for col_num, col_name in enumerate(
        df.columns,
        1
    ):

        cell = worksheet.cell(
            row=1,
            column=col_num
        )

        cell.fill = header_fill
        cell.font = header_font

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

        cell.border = thin_border


    # Format Data
    for col_idx, col in enumerate(
        df.columns,
        1
    ):

        max_len = max(
            df[col]
            .astype(str)
            .map(len)
            .max(),
            len(str(col))
        ) + 4

        col_letter = worksheet.cell(
            row=1,
            column=col_idx
        ).column_letter

        worksheet.column_dimensions[
            col_letter
        ].width = min(
            max_len,
            30
        )

        is_numeric_col = col in [
            'Area_km2',
            'Population'
        ]

        for row_idx in range(
            2,
            len(df) + 2
        ):

            cell = worksheet.cell(
                row=row_idx,
                column=col_idx
            )

            cell.border = thin_border

            if (
                is_numeric_col
                and cell.value is not None
            ):

                cell.number_format = '#,##0'

                cell.alignment = Alignment(
                    horizontal="right",
                    vertical="center"
                )

            else:

                cell.alignment = Alignment(
                    horizontal=(
                        "center"
                        if col in [
                            'No',
                            'Code',
                            'Abbreviation'
                        ]
                        else "left"
                    ),
                    vertical="center"
                )


# ==========================================
# SELESAI
# ==========================================

print("\n==========================================")
print("PROSES SELESAI")
print("==========================================")
print(f"Excel  : {excel_file}")
print(f"CSV    : {csv_file}")
print(f"Gambar : {wordcloud_file}")
print("==========================================")
