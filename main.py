import pathlib
import pandas as pd
import re
from fpdf import FPDF
from unidecode import unidecode
import pdfplumber
from docx import Document
import os


class FakultniAuditorFinal:

    def __init__(self, target_dir):
        self.target_dir = pathlib.Path(target_dir)
        self.results = []
        self.unmatched = []
        self.all_files_data = []
        self.seznam_exts = {'.pdf', '.docx', '.xlsx', '.xls', '.doc'}

        # Databáze klíčů včetně všech adres
        self.fakulty_data = {
            "Cyrilometodejska teologicka fakulta":
                r"Cyrilometodejsk[at]\s*teologick[at]\s*fakulta|Teologick[at]\s*fakulta|Univerzitni\s*22|\bcmtf\b",

            "Lekarska fakulta":
                r"Lekarsk[at]\s*fakulta|Hnevotinska\s*3|\blf\b",

            "Filozoficka fakulta":
                r"Filozofick[at]\s*fakulta|Krizkovskeho\s*10|Krizkovskeho\s*14|Vodarni\s*6|tr\.\s*Svobody\s*26|\bff\b",

            "Prirodovedecka fakulta":
                r"Prirodovedeck[at]\s*fakulta|17\.\s*listopadu\s*50a?|Slechtitelu\s*27|Slechtitelu\s*17|Slechtitelu|\bprf\b|Prirodoveda|botanick[ae]\s*zahrad[ay]|U\s*botanicke\s*zahrady\s*920|1192\/12",

            "Pedagogicka fakulta":
                r"Pedagogick[at]\s*fakulta|Zizkovo\s*nam|Purkrabska\s*2|\bpdf\b",

            "Pravnicka fakulta":
                r"Pravnick[at]\s*fakulta|17\.\s*listopadu\s*(6|8)|\bpf\b",

            "Fakulta telesne kultury":
                r"Fakulta\s*telesne\s*kultury|tr\.\s*Miru\s*(111|117)|U\s*letiste\s*32|\bftk\b|BALUO|kinantropologick",

            "Fakulta zdravotnickych ved":
                r"Fakulta\s*zdravotnickych\s*ved|Solni\s*19|Hnevotinska\s*976\/3|Hnevotinska\s*3|tr\.\s*Svobody\s*8|Stupkova\s*952\/18|\bfzv\b",

            "Koleje a menzy (SKM)":
                r"Katerinska\s*17|Smeralova\s*(12|10|8|6)|17\.\s*listopadu\s*54|tr\.\s*Miru\s*113|U\s*letiste\s*(786|827|847)|Na\s*Zakop[eě]\s*(26|4)|U\s*sportovni\s*haly\s*4|Evzena\s*Rosickeho|J\.\s*L\.\s*Fischera|Bedricha\s*Vaclavka|Vancurova\s*2|Vancurova|Kolej|koleje|\bskm\b|menza[\s_]*17\.\s*listopadu|menza[\s_]*krizkovskeho|menza[\s_]*neredin|menza[\s_]*holice|menza[\s_]*lekarsk[aé]\s*fakulta|sprava\s*koleji\s*a\s*menz",

            "Rektorat":
                r"Rektorat|rektorat|511\/8",

            "Vedeckotechnicky park (VTP UP)":
                r"Vedeckotechnicky\s*park|VTP\s*UP|Slechtitelu\s*(899|21|920\/19)|blok\s*[ABC]|Envelopa\s*Hub|1131\/8a",

            "Sport UP":
                r"Sportovni[\s_]*hala(\s*UP)?|U\s*sportovni\s*haly\s*2a?|Lodenice\s*SKUP|U\s*sportovni\s*haly\s*554|Akademik\s*sport\s*centrum",

            "Pevnost poznani":
                r"Pevnost\s*poznani|17\.\s*listopadu\s*939\/7",

            "Umelecke centrum (Konvikt)":
                r"umeleck[ée]\s*centrum|umeleck[ée]\s*centrum\s*up|umeleck[ée]\s*centrum\s*univerzity\s*palackeho|konvikt|\bucup\b|univerzitni\s*namesti\s*(3-5|3|4|5)",

            "Bazar UP":
                r"bazar[\s_]*up|bazarup|bazar",
        }

    def _get_content(self, file_path):
        ext = file_path.suffix.lower()
        try:
            if ext == ".pdf":
                with pdfplumber.open(file_path) as pdf:
                    return unidecode(" ".join([p.extract_text() or "" for p in pdf.pages[:15]]))

            elif ext == ".docx":
                doc = Document(file_path)
                return unidecode(" ".join([p.text for p in doc.paragraphs]))

        except:
            return ""

        return ""

    def run_audit(self):
        print(f"--- ANALYZA: {self.target_dir.name} ---")
        for file in self.target_dir.rglob('*'):
            if not file.is_file():
                continue

            ext = file.suffix.lower() if file.suffix else "bez pripony"
            size_mb = os.path.getsize(file) / (1024 * 1024)

            try:
                rel = file.relative_to(self.target_dir)
                root_folder = rel.parts[0] if len(rel.parts) > 1 else "Hlavni"
            except:
                root_folder = "Neznamo"

            self.all_files_data.append({
                'ext': ext,
                'size': size_mb,
                'folder': root_folder,
                'name': file.name
            })

            if ext in self.seznam_exts:
                obsah = self._get_content(file)
                stem = unidecode(file.stem.lower()).replace("_", " ")

                found_facs = []
                found_trigs = []

                for jmeno, reg in self.fakulty_data.items():
                    pattern = unidecode(reg)

                    m_obs = re.search(pattern, obsah, re.IGNORECASE)
                    m_st = re.search(pattern, stem, re.IGNORECASE)

                    if m_obs or m_st:
                        found_facs.append(jmeno)
                        t = m_obs.group(0) if m_obs else m_st.group(0)
                        found_trigs.append(t.strip())

                if found_facs:
                    self.results.append({
                        "Soubor": unidecode(file.name),
                        "Fakulty": ", ".join(set(found_facs)),
                        "Klic": ", ".join(set(found_trigs))
                    })
                else:
                    self.unmatched.append(unidecode(file.name))

    def export_pdf(self, output_name="Report4.pdf"):
        df = pd.DataFrame(self.all_files_data)

        df_summary = (
            df.groupby("ext")
            .agg(Pocet_souboru=("ext", "count"),
                 Velikost_MB=("size", "sum"))
            .reset_index()
            .sort_values(by="Pocet_souboru", ascending=False)
        )

        pdf = FPDF()

        pdf.add_page()
        pdf.set_font("Helvetica", "B", 18)
        pdf.cell(0, 15, "CELKOVY PREHLED DATASETU", align='C', new_x="LMARGIN", new_y="NEXT")
        pdf.ln(5)

        pdf.set_font("Helvetica", "B", 12)
        pdf.set_fill_color(245, 245, 245)
        pdf.cell(90, 15, f"Celkovy pocet souboru: {len(df)}", border=1, fill=True, align='C')
        pdf.cell(10, 15, "")
        pdf.cell(90, 15, f"Z toho k auditu: {len(self.results) + len(self.unmatched)}", border=1, fill=True, align='C', new_x="LMARGIN", new_y="NEXT")
        pdf.ln(10)

        pdf.set_font("Helvetica", "I", 10)
        pdf.cell(0, 8, f"Celkovy objem dat: {df['size'].sum():.2f} MB", new_x="LMARGIN", new_y="NEXT")
        pdf.ln(5)

        pdf.set_font("Helvetica", "B", 14)
        pdf.cell(0, 10, "1. Rozdeleni dle typu souboru", new_x="LMARGIN", new_y="NEXT")
        pdf.ln(5)

        # --- centrování tabulky ---
        table_width = 140
        page_width = 210
        x_offset = (page_width - table_width) / 2

        # --- HLAVIČKA ---
        pdf.set_x(x_offset)
        pdf.set_font("Helvetica", "B", 9)
        pdf.set_fill_color(230, 235, 245)

        pdf.cell(40, 8, "Typ souboru", 1, 0, 'C', True)
        pdf.cell(50, 8, "Pocet souboru", 1, 0, 'C', True)
        pdf.cell(50, 8, "Velikost (MB)", 1, 1, 'C', True)

        # --- DATA ---
        pdf.set_font("Helvetica", "", 9)

        for _, row in df_summary.iterrows():
            pdf.set_x(x_offset)
            pdf.cell(40, 8, str(row["ext"]), 1)
            pdf.cell(50, 8, f"{row['Pocet_souboru']:.0f}", 1)
            pdf.cell(50, 8, f"{row['Velikost_MB']:.2f}", 1, new_x="LMARGIN", new_y="NEXT")

        pdf.add_page()
        pdf.set_font("Helvetica", "B", 14)
        pdf.cell(0, 10, f"SEZNAM ZARAZENYCH ({len(self.results)})", align='C', new_x="LMARGIN", new_y="NEXT")
        pdf.ln(5)

        w_file, w_fac, w_key = 90, 55, 45
        pdf.set_font("Helvetica", "B", 8)
        pdf.set_fill_color(230, 235, 245)
        pdf.cell(w_file, 8, "Nazev souboru", 1, 0, 'C', True)
        pdf.cell(w_fac, 8, "Fakulta / Soucast", 1, 0, 'C', True)
        pdf.cell(w_key, 8, "Nalezeny klic", 1, 1, 'C', True)

        pdf.set_font("Helvetica", "", 7)
        for res in self.results:
            if pdf.get_y() > 260:
                pdf.add_page()

            s_x, s_y = pdf.get_x(), pdf.get_y()

            pdf.multi_cell(w_file, 5, res["Soubor"], border=1)
            y_f = pdf.get_y()

            pdf.set_xy(s_x + w_file, s_y)
            pdf.multi_cell(w_fac, 5, res["Fakulty"], border=1)
            y_fa = pdf.get_y()

            pdf.set_xy(s_x + w_file + w_fac, s_y)
            pdf.multi_cell(w_key, 5, res["Klic"], border=1)
            y_k = pdf.get_y()

            pdf.set_y(max(y_f, y_fa, y_k))

        pdf.add_page()
        pdf.set_font("Helvetica", "B", 14)
        pdf.cell(0, 10, f"SEZNAM NEZARAZENYCH ({len(self.unmatched)})", align='C', new_x="LMARGIN", new_y="NEXT")
        pdf.ln(5)

        pdf.set_font("Helvetica", "", 7)
        for name in self.unmatched:
            if pdf.get_y() > 280:
                pdf.add_page()

            pdf.multi_cell(0, 5, f"- {name}", border=0, new_x="LMARGIN", new_y="NEXT")

        pdf.output(output_name)
        print(f"Hotovo! Report: {os.path.abspath(output_name)}")


if __name__ == "__main__":
    cesta = input("Zadejte cestu ke slozce: ")
    auditor = FakultniAuditorFinal(cesta)
    auditor.run_audit()
    auditor.export_pdf()