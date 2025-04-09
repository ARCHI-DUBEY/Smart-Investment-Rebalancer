import google.generativeai as genai
import os
from dotenv import load_dotenv

load_dotenv()

api_key = os.getenv("GOOGLE_API_KEY")

# Sample User Portfolio
# user_portfolio = {
#     "AAPL": 50,  # 50 shares of Apple (Stock Ticker: AAPL)
#     "GOOGL": 30, # 30 shares of Google (Stock Ticker: GOOGL)
#     "TSLA": 20,  # 20 shares of Tesla (Stock Ticker: TSLA)
#     "VTI": 100,  # 100 shares of Vanguard Total Stock Market ETF (Ticker: VTI)
#     "AGG": 50    # 50 shares of iShares Core U.S. Aggregate Bond ETF (Ticker: AGG)
# }

try:
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel(model_name='models/gemini-1.5-pro-002')
    prompt = f"""
    Analyze the following investment portfolio:

    {user_portfolio}

    Based on this portfolio, provide investment rebalancing advice.
    Consider factors such as diversification across different asset classes (stocks and bonds) and the number of holdings.

    Suggest specific companies or sectors I should consider investing more in, and why.
    Also, suggest if there are any holdings I should consider reducing or selling, and why.

    Keep the advice concise and actionable for an individual investor.
    """

    response = model.generate_content(prompt)
    print("Rebalancing Advice:")
    print(response.text)
    print("Response from gemini-1.5-pro-002:")
    print(response.text)
except Exception as e:
    print(f"An error occurred: {e}")

# You can comment out or remove the second try block for now
# try:
#     genai.configure(api_key=api_key)
#     model_v1_0 = genai.GenerativeModel(model_name='models/gemini-1.5-pro-002')
#     response_v1_0 = model_v1_0.generate_content("What is the current price of Bitcoin?")
#     print("\nResponse from models/gemini-1.5-pro-002:")
#     print(response_v1_0.text)
# except Exception as e:
#     print(f"An error occurred with models/gemini-1.5-pro-002: {e}")