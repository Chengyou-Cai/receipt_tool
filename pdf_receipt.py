import pathlib
import zipfile
import pdfplumber
import fitz
from rich.progress import track

class PDFReceipt(object):

    def __init__(self,root=None) -> None:
        super(PDFReceipt,self).__init__()

        self.root = pathlib.Path(root)
        self._lst = list()

    def extract_zip(self):
        for path in self.root.glob(pattern=r"*通行费电子发票及详情*"):
            if path.is_dir():
                for sub_path in path.iterdir():
                    apply_zip = list(sub_path.glob(pattern="apply.zip"))[0]
                    self._lst.append(apply_zip)
                    with zipfile.ZipFile(apply_zip) as zip:
                        zip.extractall(path="./_pdf")
        print(f"Number of Zipfile : {len(self._lst)}") # 2+5+6+7+17+6=43

    def process_pdf(self):
        for path in track(pathlib.Path("./_pdf").glob("*.pdf")):
            numb = None

            with pdfplumber.open(path.as_posix()) as pdf:
                page = pdf.pages[0]
                numb = page.crop((473, 34, 516, 49)).extract_text()
            
            with fitz.open(path.as_posix()) as pdf:
                page = pdf.load_page(0)
                crop = page.get_pixmap(matrix=fitz.Matrix(1,1),clip=fitz.Rect(page.rect))
                crop.save(f"_pic/{numb}.jpg")

if __name__ == "__main__":
    receipt = PDFReceipt(r"C:\Users\cheng\Desktop\workspace\发票")
    # receipt.extract_zip()
    receipt.process_pdf()
