import tkinter as tk
from tkinter import ttk, messagebox
from tkinter import filedialog
from instruction_creator import create_instruction_docx

def main():
    def select_text_file():
        text_path.set(filedialog.askopenfilename(title="手順テキストファイルを選択"))
    
    def select_image_directory():
        img_dir.set(filedialog.askdirectory(title="画像ディレクトリを選択"))
    
    def select_output_path():
        output_path.set(filedialog.asksaveasfilename(defaultextension='.docx', title="保存先を指定"))
    
    def execute_creation():
        # チェックポイント: 全てのパスが選択されているか確認
        if not text_path.get():
            messagebox.showerror("エラー", "テキストファイルが選択されていません。")
            return
        if not img_dir.get():
            messagebox.showerror("エラー", "画像ディレクトリが選択されていません。")
            return
        if not output_path.get():
            messagebox.showerror("エラー", "出力先のパスが選択されていません。")
            return
        
        # 全て選択されていれば、処理を開始
        create_instruction_docx(text_path.get(), img_dir.get(), output_path.get())
    
    root = tk.Tk()
    root.title("手順書作成アプリ")
    root.geometry("700x500")
    root.configure(bg="#1c2833")

    style = ttk.Style(root)
    style.theme_use('clam')
    
    style.configure('Title.TLabel', font=('Arial', 34, 'bold'), foreground="#ffffff", background="#1c2833")
    style.configure('Desc.TLabel', font=('Arial', 16), foreground="#d0d0d0", background="#1c2833")
    style.configure('TButton', font=('Arial', 14, 'bold'), foreground='#1c2833', background='#e67e22', borderwidth=0, focuscolor='none', padding=8)
    style.map('TButton', background=[('active', '#f39c12'), ('pressed', '#d35400')], foreground=[('active', '#ffffff')])

    # 変数
    text_path = tk.StringVar()
    img_dir = tk.StringVar()
    output_path = tk.StringVar()

    # タイトル
    title_label = ttk.Label(root, text="手順書作成アプリ", style='Title.TLabel', anchor=tk.CENTER)
    title_label.pack(pady=(30, 10))

    # 説明文
    description = "このアプリではテキストと画像を使用して\n手順書を簡単に作成できます。"
    desc_label = ttk.Label(root, text=description, style='Desc.TLabel', justify=tk.CENTER, anchor=tk.CENTER)
    desc_label.pack(pady=(0, 40))

    # ファイル選択ボタンとラベル
    button_frame = ttk.Frame(root)
    button_frame.pack(pady=10)

    ttk.Button(button_frame, text="テキストファイルを選択", command=select_text_file).grid(row=0, column=0, padx=5, pady=5)
    ttk.Label(button_frame, textvariable=text_path, style='Desc.TLabel').grid(row=0, column=1, padx=5, pady=5)
    
    ttk.Button(button_frame, text="画像ディレクトリを選択", command=select_image_directory).grid(row=1, column=0, padx=5, pady=5)
    ttk.Label(button_frame, textvariable=img_dir, style='Desc.TLabel').grid(row=1, column=1, padx=5, pady=5)
    
    ttk.Button(button_frame, text="出力パスを選択", command=select_output_path).grid(row=2, column=0, padx=5, pady=5)
    ttk.Label(button_frame, textvariable=output_path, style='Desc.TLabel').grid(row=2, column=1, padx=5, pady=5)

    # 作成ボタン
    create_button = ttk.Button(root, text="作成", command=execute_creation)
    create_button.pack(ipadx=30, ipady=15, pady=20)

    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - (root.winfo_width() // 2)
    y = (root.winfo_screenheight() // 2) - (root.winfo_height() // 2)
    root.geometry(f'+{x}+{y}')

    root.mainloop()

if __name__ == '__main__':
    main()