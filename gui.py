import tkinter as tk
from tkinter import scrolledtext, messagebox, filedialog
from PIL import Image
import pytesseract
import json
import google.generativeai as genai
import os
from dotenv import load_dotenv

load_dotenv()
api_key = os.getenv("GOOGLE_API_KEY")
genai.configure(api_key=api_key)
model = genai.GenerativeModel(model_name='models/gemini-1.5-pro-002')

def analyze_portfolio():
    import pytesseract
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'  # Adjust if your path is different
    input_method = input_method_var.get()
    result_text.delete("1.0", tk.END)

    if input_method == "text":
        portfolio_text = portfolio_input.get("1.0", tk.END)
        try:
            portfolio = json.loads(portfolio_text)
            advice = get_rebalancing_advice(portfolio)
            result_text.insert(tk.END, advice)
        except json.JSONDecodeError:
            messagebox.showerror("Error", "Invalid JSON format.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")
    elif input_method == "image":
        file_path = filedialog.askopenfilename(
            title="Select Portfolio Image",
            filetypes=[("Image files", "*.png;*.jpg;*.jpeg")]
        )
        if file_path:
            try:
                img = Image.open(file_path)
                extracted_text = pytesseract.image_to_string(img)
                result_text.insert(tk.END, f"Extracted Text:\n{extracted_text}\n\nAttempting to parse portfolio...\n")
                portfolio = parse_portfolio_from_text(extracted_text)
                if portfolio:
                    advice = get_rebalancing_advice(portfolio)
                    result_text.insert(tk.END, f"\nRebalancing Advice:\n{advice}")
                else:
                    messagebox.showerror("Error", "Could not parse portfolio information from the image.")
            except FileNotFoundError:
                messagebox.showerror("Error", "Image file not found.")
            except Exception as e:
                messagebox.showerror("Error", f"Error processing image: {e}")
    elif input_method == "excel":
        messagebox.showinfo("Info", "Excel analysis is not yet implemented.") # Placeholder
    elif input_method == "pdf":
        messagebox.showinfo("Info", "PDF analysis is not yet implemented.")   # Placeholder

def get_rebalancing_advice(portfolio):
    prompt = f"""
    Analyze the following investment portfolio:

    {portfolio}

    Provide investment rebalancing advice... (rest of your prompt)
    """
    response = model.generate_content(prompt)
    return response.text

def parse_portfolio_from_text(text):
    portfolio = {}
    lines = text.split('\n')
    for line in lines:
        if "Stocks" in line:
            parts = line.split("Stocks")
            if len(parts) == 2:
                company_info = parts[0].strip()
                quantity_value = parts[1].strip().split()
                if len(quantity_value) >= 1:
                    try:
                        shares = int(quantity_value[0])
                        # We need to make an assumption or have a better way to get the ticker.
                        # For now, let's take the first word of the company name as a potential ticker.
                        ticker_parts = company_info.split()
                        if ticker_parts:
                            ticker = ticker_parts[0].strip().replace('.', '').replace(',', '') # Basic cleaning
                            portfolio[ticker] = shares
                    except ValueError:
                        pass
    return portfolio
window = tk.Tk()
window.title("Smart AI Investment Rebalancer")

input_method_var = tk.StringVar(window)
input_method_var.set("text") # Default input method

input_method_label = tk.Label(window, text="Input Method:")
input_method_label.pack()

input_method_menu = tk.OptionMenu(window, input_method_var, "text", "image", "excel", "pdf")
input_method_menu.pack()

portfolio_label = tk.Label(window, text="Enter Portfolio (JSON for text):")
portfolio_label.pack()

portfolio_input = scrolledtext.ScrolledText(window, height=5, width=50)
portfolio_input.pack()

browse_button = tk.Button(window, text="Browse Image", command=analyze_portfolio)
browse_button.pack()

analyze_button = tk.Button(window, text="Analyze Portfolio", command=analyze_portfolio)
analyze_button.pack()

result_text = scrolledtext.ScrolledText(window, height=15, width=60)
result_text.pack()

window.mainloop()