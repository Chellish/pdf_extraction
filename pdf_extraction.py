import os
from tkinter import Tk, Label, Button, filedialog, messagebox
from PyPDF2 import PdfReader, PdfWriter
import pandas as pd

class PDFExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF抽出ツール")
        
        self.pdf_path = None
        self.csv_path = None
        
        Label(root, text="PDFファイル:").grid(row=0, column=0, padx=10, pady=5, sticky="w")
        Label(root, text="CSVファイル:").grid(row=1, column=0, padx=10, pady=5, sticky="w")
        
        self.pdf_label = Label(root, text="未選択", width=40, anchor="w", bg="white", relief="sunken")
        self.pdf_label.grid(row=0, column=1, padx=10, pady=5)
        
        self.csv_label = Label(root, text="未選択", width=40, anchor="w", bg="white", relief="sunken")
        self.csv_label.grid(row=1, column=1, padx=10, pady=5)
        
        Button(root, text="PDFを選択", command=self.select_pdf).grid(row=0, column=2, padx=10, pady=5)
        Button(root, text="CSVを選択", command=self.select_csv).grid(row=1, column=2, padx=10, pady=5)
        
        Button(root, text="抽出を実行", command=self.extract_pages).grid(row=2, column=0, columnspan=3, pady=10)

    def select_pdf(self):
        self.pdf_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        self.pdf_label.config(text=self.pdf_path if self.pdf_path else "未選択")
    
    def select_csv(self):
        self.csv_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        self.csv_label.config(text=self.csv_path if self.csv_path else "未選択")
    
    def extract_pages(self):
        if not self.pdf_path or not self.csv_path:
            messagebox.showerror("エラー", "PDFまたはCSVが選択されていません。")
            return
        
        try:
            # CSVファイルを読み込み
            conditions = pd.read_csv(self.csv_path, sep=",", header=None, names=["filename", "start_page", "end_page"])
            
            # 欠損値を削除
            conditions = conditions.dropna()
            conditions["start_page"] = conditions["start_page"].astype(int)
            conditions["end_page"]   = conditions["end_page"].astype(int)
            
            # PDFを読み込む
            reader       = PdfReader(self.pdf_path)
            total_pages  = len(reader.pages)
            input_folder = os.path.dirname(self.pdf_path)
            tmp_folder   = os.path.join(input_folder, "tmp")
            
            # tmpフォルダを作成（存在しない場合）
            if not os.path.exists(tmp_folder):
                os.makedirs(tmp_folder)
            
            # 抽出処理
            unprocessed_files = []
            for _, row in conditions.iterrows():
                output_pdf_path = os.path.join(tmp_folder, f"{row['filename']}.pdf")  # ファイル名に .pdf を追加
                writer = PdfWriter()
                # 0始まりに調整
                start_page = row["start_page"] - 1  
                end_page = row["end_page"] - 1
                
                if start_page >= total_pages:
                    # 開始ページがPDFの総ページ数を超える場合
                    unprocessed_files.append(f"{row['filename']} (開始ページ超過: {row['start_page']}～{row['end_page']})")
                    continue
                
                # 実際に抽出可能な範囲を調整
                actual_end_page = min(end_page, total_pages - 1)
                
                for page_num in range(start_page, actual_end_page + 1):
                    writer.add_page(reader.pages[page_num])
                
                with open(output_pdf_path, "wb") as output_file:
                    writer.write(output_file)
                
                print(f"抽出完了: {output_pdf_path}")
            
            # 処理できなかったファイルをGUIで表示
            if unprocessed_files:
                error_message = "\n".join(unprocessed_files)
                messagebox.showwarning("未処理のファイル", f"以下の条件を満たす抽出はできませんでした:\n{error_message}")
            else:
                messagebox.showinfo("成功", "すべての抽出条件を満たしました。")
            
        except Exception as e:
            messagebox.showerror("エラー", f"処理中にエラーが発生しました:\n{e}")

# 実行
if __name__ == "__main__":
    root = Tk()
    app  = PDFExtractorApp(root)
    root.mainloop()
