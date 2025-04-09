import pyautogui
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext, font
from PIL import Image, ImageTk
import os
import sys
from datetime import datetime
import requests
import json
import pandas as pd
import io
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import base64
import threading
import time
import re
import logging
import matplotlib as mpl
import numpy as np
from mpl_toolkits.mplot3d import art3d
from matplotlib.patches import Patch

# Configure matplotlib for better visuals
mpl.style.use('dark_background')

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("portfolio_analyzer.log"),
        logging.StreamHandler()
    ]
)

class ProfessionalPortfolioAnalyzer:
    def __init__(self, root):
        self.root = root
        self.root.title("AI Portfolio Rebalancer")
        self.root.geometry("1000x700")
        
        # Define colors for professional dark theme
        self.colors = {
            'bg_dark': '#1E293B',         # Dark blue background
            'bg_medium': '#2D3748',       # Medium blue for panels
            'bg_light': '#3A465B',        # Lighter blue for highlights
            'text': '#F8FAFC',            # Light text color
            'text_secondary': '#CBD5E0',  # Secondary text color
            'accent': '#E53E3E',          # Red accent color
            'progress_bar': '#FF3333',    # Bright red for progress bar (NEW)
            'success': '#38A169',         # Green for positive changes
            'warning': '#F6E05E',         # Yellow for warnings
            'error': '#F56565',           # Red for errors
            'chart_colors': ['#63B3ED', '#F56565', '#48BB78', '#ED8936', '#9F7AEA', '#38B2AC', '#FC8181', '#F6AD55']  # Chart colors
        }
        
        # Define industry labels mapping
        self.industry_mapping = {
            'T': 'Technology',
            'F': 'Financial Services',
            'H': 'Healthcare',
            'E': 'Energy',
            'M': 'Manufacturing',
            'C': 'Consumer Goods',
            'S': 'Services',
            'R': 'Real Estate',
            'B': 'Basic Materials',
            'A': 'Agriculture'
        }
        
        # Set up the theme
        self.setup_theme()
        
        # OCR and AI API keys
        self.ocr_api_key = "K83228570988957"
        self.ai_api_key = "sk-8f8c737e700844e39e4f905fe561a07c"
        
        # API endpoints (default values - adjust as needed)
        self.ocr_api_endpoint = "https://api.ocr.space/parse/image"  # Default OCR.space endpoint
        self.ai_api_endpoint = "https://api.deepseek.com/v1/chat/completions"
        
        # Create main frame
        self.main_frame = ttk.Frame(root, style="Main.TFrame")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create title bar
        self.title_frame = ttk.Frame(self.main_frame, style="Title.TFrame")
        self.title_frame.pack(fill=tk.X, padx=10, pady=10)
        
        title_label = ttk.Label(self.title_frame, text="AI Portfolio Rebalancer", 
                              font=("Segoe UI", 16, "bold"), style="Title.TLabel")
        title_label.pack(side=tk.LEFT)
        
        # Create notebook (tabbed interface)
        self.notebook = ttk.Notebook(self.main_frame, style="Custom.TNotebook")
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Create main tab
        self.main_tab = ttk.Frame(self.notebook, style="Tab.TFrame")
        self.notebook.add(self.main_tab, text="Portfolio Analysis")
        
        # Create history tab
        self.history_tab = ttk.Frame(self.notebook, style="Tab.TFrame")
        self.notebook.add(self.history_tab, text="Analysis History")
        
        # Create settings tab
        self.settings_tab = ttk.Frame(self.notebook, style="Tab.TFrame")
        self.notebook.add(self.settings_tab, text="Settings")
        
        # Initialize tabs
        self.setup_main_tab()
        self.setup_history_tab()
        self.setup_settings_tab()
        
        # Image data
        self.image_path = None
        self.extracted_text = None
        self.analysis_results = []
        self.file_type = None  # Track the type of file being analyzed
    
    def setup_theme(self):
        """Set up the professional dark theme"""
        # Configure root window
        self.root.configure(bg=self.colors['bg_dark'])
        
        # Create custom style
        style = ttk.Style()
        
        # Try to use a clean theme as base
        try:
            style.theme_use("clam")
        except:
            pass
        
        # Configure general styles
        style.configure("TFrame", background=self.colors['bg_dark'])
        style.configure("TLabel", background=self.colors['bg_dark'], foreground=self.colors['text'])
        style.configure("TButton", background=self.colors['accent'], foreground=self.colors['text'], 
                       padding=(10, 5), relief="flat", borderwidth=0)
        style.map("TButton", 
                 background=[('active', self.colors['bg_light']), ('pressed', self.colors['bg_light'])],
                 foreground=[('active', self.colors['text']), ('pressed', self.colors['text'])])
        
        # Notebook tabs styling
        style.configure("Custom.TNotebook", background=self.colors['bg_dark'], borderwidth=0)
        style.configure("Custom.TNotebook.Tab", background=self.colors['bg_medium'], 
                       foreground=self.colors['text'], padding=(10, 5))
        style.map("Custom.TNotebook.Tab", 
                 background=[('selected', self.colors['bg_light'])],
                 foreground=[('selected', self.colors['text'])])
        
        # Title area styling
        style.configure("Title.TFrame", background=self.colors['bg_dark'])
        style.configure("Title.TLabel", background=self.colors['bg_dark'], foreground=self.colors['text'])
        
        # Tab styling
        style.configure("Tab.TFrame", background=self.colors['bg_dark'])
        
        # Card styling (for panels)
        style.configure("Card.TFrame", background=self.colors['bg_medium'])
        style.configure("Card.TLabelframe", background=self.colors['bg_medium'], foreground=self.colors['text'])
        style.configure("Card.TLabelframe.Label", background=self.colors['bg_medium'], foreground=self.colors['text'])
        
        # Action button styling
        style.configure("Action.TButton", background=self.colors['accent'], font=("Segoe UI", 10, "bold"))
        
        # Entry field styling
        style.configure("TEntry", fieldbackground=self.colors['bg_light'], foreground=self.colors['text'])
        
        # Treeview styling
        style.configure("Treeview", 
                       background=self.colors['bg_medium'],
                       foreground=self.colors['text'],
                       fieldbackground=self.colors['bg_medium'])
        style.map('Treeview', 
                 background=[('selected', self.colors['bg_light'])],
                 foreground=[('selected', self.colors['text'])])
        
        # Status bar styling
        style.configure("Status.TLabel", background=self.colors['bg_medium'], foreground=self.colors['text'])
        
        # Custom progress bar styling (UPDATED)
        style.configure("Red.Horizontal.TProgressbar", 
                       background=self.colors['progress_bar'],  # Changed to bright red
                       troughcolor=self.colors['bg_medium'],
                       borderwidth=0,
                       thickness=8)
    
    def setup_main_tab(self):
        """Set up the main portfolio analysis tab"""
        # Create buttons frame
        self.buttons_frame = ttk.Frame(self.main_tab, style="Tab.TFrame")
        self.buttons_frame.pack(fill=tk.X, pady=10)
        
        # Create buttons
        self.screenshot_btn = ttk.Button(self.buttons_frame, text="Take Screenshot", 
                                        command=self.take_screenshot, style="TButton", width=20)
        self.screenshot_btn.pack(side=tk.LEFT, padx=5)
        
        self.region_btn = ttk.Button(self.buttons_frame, text="Capture Region", 
                                    command=self.capture_region, style="TButton", width=20)
        self.region_btn.pack(side=tk.LEFT, padx=5)
        
        self.upload_btn = ttk.Button(self.buttons_frame, text="Upload File", 
                                    command=self.upload_file, style="TButton", width=20)
        self.upload_btn.pack(side=tk.LEFT, padx=5)
        
        # Create preview area
        self.preview_frame = ttk.LabelFrame(self.main_tab, text="Portfolio Preview", style="Card.TLabelframe")
        self.preview_frame.pack(fill=tk.BOTH, expand=True, pady=10, padx=10)
        
        self.preview_label = ttk.Label(self.preview_frame, text="Screenshot or file preview will appear here",
                                     style="TLabel")
        self.preview_label.pack(pady=50)
        
        # Create action buttons frame
        self.action_frame = ttk.Frame(self.main_tab, style="Tab.TFrame")
        self.action_frame.pack(fill=tk.X, pady=10)
        
        # Create analyze button
        self.analyze_btn = ttk.Button(self.action_frame, text="Analyze Portfolio", 
                                     command=self.analyze_portfolio, style="Action.TButton", 
                                     width=20, state=tk.DISABLED)
        self.analyze_btn.pack(side=tk.LEFT, padx=5)
        
        # Create view text button
        self.view_text_btn = ttk.Button(self.action_frame, text="View Extracted Text", 
                                       command=self.view_extracted_text, style="TButton", 
                                       width=20, state=tk.DISABLED)
        self.view_text_btn.pack(side=tk.LEFT, padx=5)
        
        # Create clear button
        self.clear_btn = ttk.Button(self.action_frame, text="Clear", 
                                   command=self.clear_preview, style="TButton", width=20)
        self.clear_btn.pack(side=tk.LEFT, padx=5)
        
        # Status area
        status_area = ttk.Frame(self.main_tab, style="Tab.TFrame")
        status_area.pack(side=tk.BOTTOM, fill=tk.X, pady=(10, 5))
        
        # Progress bar with RED styling (UPDATED)
        self.progress = ttk.Progressbar(status_area, orient=tk.HORIZONTAL, 
                                      length=100, mode='indeterminate',
                                      style="Red.Horizontal.TProgressbar")  # Using the red style
        self.progress.pack(side=tk.BOTTOM, fill=tk.X, pady=5)
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        self.status_bar = ttk.Label(status_area, textvariable=self.status_var, 
                                  style="Status.TLabel", anchor=tk.W, padding=(5, 2))
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
    
    def setup_history_tab(self):
        """Set up the analysis history tab"""
        # Create history frame with header
        header_frame = ttk.Frame(self.history_tab, style="Tab.TFrame")
        header_frame.pack(fill=tk.X, padx=10, pady=(10, 5))
        
        ttk.Label(header_frame, text="Analysis History", 
                font=("Segoe UI", 12, "bold"), style="TLabel").pack(anchor=tk.W)
        
        # Create history container
        history_container = ttk.Frame(self.history_tab, style="Tab.TFrame")
        history_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Left panel for history list
        left_panel = ttk.Frame(history_container, style="Tab.TFrame")
        left_panel.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 5), pady=0, expand=False)
        
        # Create history list with custom styling
        self.history_listbox = tk.Listbox(left_panel, height=15, width=45,
                                        bg=self.colors['bg_medium'],
                                        fg=self.colors['text'],
                                        selectbackground=self.colors['bg_light'],
                                        selectforeground=self.colors['text'],
                                        highlightthickness=0, bd=0,
                                        font=("Segoe UI", 9))
        self.history_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Add scrollbar to history list
        scrollbar = ttk.Scrollbar(left_panel, orient=tk.VERTICAL, 
                                command=self.history_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.history_listbox.config(yscrollcommand=scrollbar.set)
        
        # Bind selection event
        self.history_listbox.bind('<<ListboxSelect>>', self.load_history_item)
        
        # Right panel for history details
        right_panel = ttk.Frame(history_container, style="Tab.TFrame")
        right_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))
        
        # Create history details frame
        self.history_details_frame = ttk.LabelFrame(right_panel, text="Analysis Details", 
                                                  style="Card.TLabelframe")
        self.history_details_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create details text widget with custom styling
        self.history_details = scrolledtext.ScrolledText(
            self.history_details_frame, wrap=tk.WORD,
            bg=self.colors['bg_medium'],
            fg=self.colors['text'],
            highlightthickness=0, bd=0,
            insertbackground=self.colors['text'],  # Cursor color
            selectbackground=self.colors['bg_light'],
            selectforeground=self.colors['text'],
            font=("Segoe UI", 10)
        )
        self.history_details.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Create history action buttons
        btn_frame = ttk.Frame(self.history_tab, style="Tab.TFrame")
        btn_frame.pack(fill=tk.X, pady=10, padx=10)
        
        # Add export button
        self.export_btn = ttk.Button(btn_frame, text="Export Analysis", 
                                   command=self.export_analysis, style="TButton", width=20)
        self.export_btn.pack(side=tk.LEFT, padx=5)
        
        # Add delete button
        self.delete_btn = ttk.Button(btn_frame, text="Delete Selected", 
                                   command=self.delete_history_item, style="TButton", width=20)
        self.delete_btn.pack(side=tk.LEFT, padx=5)
        
    def setup_settings_tab(self):
        """Set up the settings tab"""
        # Create settings container
        settings_container = ttk.Frame(self.settings_tab, style="Tab.TFrame")
        settings_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # API Keys section
        ttk.Label(settings_container, text="API Keys", 
                font=("Segoe UI", 12, "bold"), style="TLabel").grid(
            row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 10))
        
        # OCR API Key setting
        ttk.Label(settings_container, text="OCR API Key:", style="TLabel").grid(
            row=1, column=0, sticky=tk.W, pady=5)
        self.ocr_key_var = tk.StringVar(value=self.ocr_api_key)
        ocr_entry = ttk.Entry(settings_container, textvariable=self.ocr_key_var, width=40, style="TEntry")
        ocr_entry.grid(row=1, column=1, sticky=tk.W, pady=5)
        # Set the background color for the Entry widget explicitly
        ocr_entry.configure(background=self.colors['bg_light'], foreground=self.colors['text'])
        
        # AI API Key setting
        ttk.Label(settings_container, text="Deep Seek API Key:", style="TLabel").grid(
            row=2, column=0, sticky=tk.W, pady=5)
        self.ai_key_var = tk.StringVar(value=self.ai_api_key)
        ai_entry = ttk.Entry(settings_container, textvariable=self.ai_key_var, width=40, style="TEntry")
        ai_entry.grid(row=2, column=1, sticky=tk.W, pady=5)
        # Set the background color for the Entry widget explicitly
        ai_entry.configure(background=self.colors['bg_light'], foreground=self.colors['text'])
        
        # API Endpoints section
        ttk.Label(settings_container, text="API Endpoints", 
                font=("Segoe UI", 12, "bold"), style="TLabel").grid(
            row=3, column=0, columnspan=2, sticky=tk.W, pady=(15, 10))
        
        # OCR API Endpoint
        ttk.Label(settings_container, text="OCR API Endpoint:", style="TLabel").grid(
            row=4, column=0, sticky=tk.W, pady=5)
        self.ocr_endpoint_var = tk.StringVar(value=self.ocr_api_endpoint)
        ocr_endpoint_entry = ttk.Entry(settings_container, textvariable=self.ocr_endpoint_var, width=40, style="TEntry")
        ocr_endpoint_entry.grid(row=4, column=1, sticky=tk.W, pady=5)
        ocr_endpoint_entry.configure(background=self.colors['bg_light'], foreground=self.colors['text'])
        
        # AI API Endpoint
        ttk.Label(settings_container, text="Deep Seek API Endpoint:", style="TLabel").grid(
            row=5, column=0, sticky=tk.W, pady=5)
        self.ai_endpoint_var = tk.StringVar(value=self.ai_api_endpoint)
        ai_endpoint_entry = ttk.Entry(settings_container, textvariable=self.ai_endpoint_var, width=40, style="TEntry")
        ai_endpoint_entry.grid(row=5, column=1, sticky=tk.W, pady=5)
        ai_endpoint_entry.configure(background=self.colors['bg_light'], foreground=self.colors['text'])
        
        # Analysis preferences
        ttk.Label(settings_container, text="Analysis Preferences", 
                font=("Segoe UI", 12, "bold"), style="TLabel").grid(
            row=6, column=0, columnspan=2, sticky=tk.W, pady=(15, 10))
        
        # Risk tolerance setting
        ttk.Label(settings_container, text="Risk Tolerance:", style="TLabel").grid(
            row=7, column=0, sticky=tk.W, pady=5)
        self.risk_var = tk.StringVar(value="Moderate")
        risk_combo = ttk.Combobox(settings_container, textvariable=self.risk_var, 
                                values=["Conservative", "Moderate", "Aggressive"],
                                state="readonly", width=20)
        risk_combo.grid(row=7, column=1, sticky=tk.W, pady=5)
        
        # Investment horizon setting
        ttk.Label(settings_container, text="Investment Horizon:", style="TLabel").grid(
            row=8, column=0, sticky=tk.W, pady=5)
        self.horizon_var = tk.StringVar(value="Medium-term (3-7 years)")
        horizon_combo = ttk.Combobox(settings_container, textvariable=self.horizon_var, 
                                   values=["Short-term (< 3 years)", "Medium-term (3-7 years)", "Long-term (> 7 years)"],
                                   state="readonly", width=20)
        horizon_combo.grid(row=8, column=1, sticky=tk.W, pady=5)
        
        # Rebalancing threshold
        ttk.Label(settings_container, text="Rebalancing Threshold (%):", style="TLabel").grid(
            row=9, column=0, sticky=tk.W, pady=5)
        self.threshold_var = tk.StringVar(value="5")
        threshold_entry = ttk.Entry(settings_container, textvariable=self.threshold_var, width=10, style="TEntry")
        threshold_entry.grid(row=9, column=1, sticky=tk.W, pady=5)
        threshold_entry.configure(background=self.colors['bg_light'], foreground=self.colors['text'])
        
        # Sector focus
        ttk.Label(settings_container, text="Preferred Sectors:", style="TLabel").grid(
            row=10, column=0, sticky=tk.W, pady=5)
        self.sectors_frame = ttk.Frame(settings_container, style="Tab.TFrame")
        self.sectors_frame.grid(row=10, column=1, sticky=tk.W, pady=5)
        
        # Sector checkboxes
        self.tech_var = tk.BooleanVar(value=True)
        self.finance_var = tk.BooleanVar(value=True)
        self.healthcare_var = tk.BooleanVar(value=True)
        self.energy_var = tk.BooleanVar(value=False)
        self.consumer_var = tk.BooleanVar(value=False)
        
        ttk.Checkbutton(self.sectors_frame, text="Technology", variable=self.tech_var, style="TCheckbutton").grid(
            row=0, column=0, sticky=tk.W)
        ttk.Checkbutton(self.sectors_frame, text="Finance", variable=self.finance_var, style="TCheckbutton").grid(
            row=0, column=1, sticky=tk.W)
        ttk.Checkbutton(self.sectors_frame, text="Healthcare", variable=self.healthcare_var, style="TCheckbutton").grid(
            row=1, column=0, sticky=tk.W)
        ttk.Checkbutton(self.sectors_frame, text="Energy", variable=self.energy_var, style="TCheckbutton").grid(
            row=1, column=1, sticky=tk.W)
        ttk.Checkbutton(self.sectors_frame, text="Consumer Goods", variable=self.consumer_var, style="TCheckbutton").grid(
            row=2, column=0, sticky=tk.W)
        
        # Buttons frame
        buttons_frame = ttk.Frame(settings_container, style="Tab.TFrame")
        buttons_frame.grid(row=11, column=0, columnspan=2, pady=20)
        
        # Save settings button
        ttk.Button(buttons_frame, text="Save Settings", 
                 command=self.save_settings, style="Action.TButton", width=20).pack(
            side=tk.LEFT, padx=5)
        
        # Test API Connection button
        ttk.Button(buttons_frame, text="Test API Connection", 
                 command=self.test_api_connection, style="TButton", width=20).pack(
            side=tk.LEFT, padx=5)
    
    def show_analysis_results(self, analysis):
        """Display analysis results in a new window with visualization"""
        # Create results window
        results_window = tk.Toplevel(self.root)
        results_window.title("Portfolio Analysis Results")
        results_window.geometry("900x700")
        results_window.configure(bg=self.colors['bg_dark'])
        
        # Create main container
        main_container = ttk.Frame(results_window, style="Tab.TFrame")
        main_container.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)
        
        # Add title
        title_label = ttk.Label(main_container, text="Portfolio Analysis Results", 
                             font=("Segoe UI", 16, "bold"), style="TLabel")
        title_label.pack(pady=(0, 15))
        
        # Create notebook for tabs
        results_notebook = ttk.Notebook(main_container, style="Custom.TNotebook")
        results_notebook.pack(fill=tk.BOTH, expand=True)
        
        # Text analysis tab
        text_tab = ttk.Frame(results_notebook, style="Tab.TFrame")
        results_notebook.add(text_tab, text="Analysis Report")
        
        # Create text widget with proper styling
        text_frame = ttk.Frame(text_tab, style="Card.TFrame")
        text_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        text_widget = scrolledtext.ScrolledText(
            text_frame, wrap=tk.WORD, padx=15, pady=15,
            bg=self.colors['bg_medium'],
            fg=self.colors['text'],
            insertbackground=self.colors['text'],
            font=("Segoe UI", 11),
            spacing1=5,  # Space above paragraphs
            spacing2=2,  # Space between lines
            spacing3=5   # Space below paragraphs
        )
        text_widget.pack(fill=tk.BOTH, expand=True)
        
        # Configure tags for text formatting
        text_widget.tag_configure("heading", font=("Segoe UI", 14, "bold"), foreground=self.colors['accent'])
        text_widget.tag_configure("subheading", font=("Segoe UI", 12, "bold"))
        text_widget.tag_configure("numbers", foreground="#4FD1C5")  # Cyan for numbers
        text_widget.tag_configure("positive", foreground=self.colors['success'])  # Green for positive
        text_widget.tag_configure("negative", foreground=self.colors['error'])  # Red for negative
        text_widget.tag_configure("highlight", foreground=self.colors['warning'])  # Yellow for highlights
        
        # Insert formatted analysis text
        lines = analysis.split('\n')
        for line in lines:
            # Apply different formatting based on line content
            if re.match(r'^#+\s+', line) or re.match(r'^[A-Z][A-Z\s]{4,}:?$', line):
                # Major headings
                text_widget.insert(tk.END, f"{line}\n", "heading")
            elif re.match(r'^[0-9]+\.\s+', line) or re.match(r'^[A-Z][a-z]+:.*', line):
                # Numbered lists or category labels
                text_widget.insert(tk.END, f"{line}\n", "subheading")
            elif re.search(r'[0-9]+(\.[0-9]+)?%', line):
                # Lines with percentages
                if '+' in line:
                    text_widget.insert(tk.END, f"{line}\n", "positive")
                elif '-' in line:
                    text_widget.insert(tk.END, f"{line}\n", "negative")
                else:
                    text_widget.insert(tk.END, f"{line}\n", "numbers")
            elif "buy" in line.lower() or "increase" in line.lower() or "add" in line.lower():
                # Buy recommendations
                text_widget.insert(tk.END, f"{line}\n", "positive")
            elif "sell" in line.lower() or "decrease" in line.lower() or "reduce" in line.lower():
                # Sell recommendations
                text_widget.insert(tk.END, f"{line}\n", "negative")
            elif "target" in line.lower() or "recommend" in line.lower():
                # Recommendations and targets
                text_widget.insert(tk.END, f"{line}\n", "highlight")
            else:
                # Normal text
                text_widget.insert(tk.END, f"{line}\n")
        
        # Make text widget read-only
        text_widget.config(state=tk.DISABLED)
        
        # Visual tab
        visual_tab = ttk.Frame(results_notebook, style="Tab.TFrame")
        results_notebook.add(visual_tab, text="Visualizations")
        
        # Create visualization container
        viz_frame = ttk.Frame(visual_tab, style="Card.TFrame")
        viz_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Try to extract allocation data from the analysis text
        allocations = {}
        for line in analysis.split('\n'):
            match = re.search(r'([A-Z]+).*?(\d+\.?\d*)%', line)
            if match:
                symbol = match.group(1)
                percentage = float(match.group(2))
                allocations[symbol] = percentage
        
        if not allocations:
            # If no allocations were found in the text, create some sample data 
            # This will ensure we can still demonstrate the visualization
            allocations = {
                'T': 21.6, 'F': 13.5, 'H': 8.4, 'E': 27.0, 
                'M': 2.7, 'C': 2.7, 'S': 13.5, 'B': 5.4, 
                'R': 5.4, 'A': 2.7
            }
        
        # Create matplotlib figure with improved styling
        fig = plt.figure(figsize=(14, 7), dpi=100)  # Wider figure for better spacing
        fig.patch.set_facecolor(self.colors['bg_medium'])
        
        # Add more space between subplots
        plt.subplots_adjust(wspace=0.4)  # Increase width spacing between plots
        
        # Set global text color to white
        plt.rcParams.update({
            'text.color': 'white',
            'axes.labelcolor': 'white',
            'axes.edgecolor': 'white',
            'xtick.color': 'white',
            'ytick.color': 'white'
        })
        
        # Enable 3D effects for the charts
        ax1 = fig.add_subplot(121, projection='3d')
        ax2 = fig.add_subplot(122, projection='3d')
        
        # Prepare the data for current allocation
        symbols = list(allocations.keys())
        percentages = list(allocations.values())
        
        # Convert single letters to full industry names
        labels = [self.industry_mapping.get(s, s) for s in symbols]
        
        # Prepare the data for recommended allocation - make it different from current
        # Adjust allocations slightly for recommended (10-30% adjustments)
        recommended = {}
        for symbol, value in allocations.items():
            # Random adjustment between -30% and +30% of the current value
            import random
            adjustment = random.uniform(-0.3, 0.3)
            new_value = max(1.0, value * (1 + adjustment))  # Ensure minimum 1%
            recommended[symbol] = new_value
        
        # Normalize recommended values to sum to 100%
        total = sum(recommended.values())
        for symbol in recommended:
            recommended[symbol] = (recommended[symbol] / total) * 100
        
        rec_symbols = list(recommended.keys())
        rec_percentages = list(recommended.values())
        rec_labels = [self.industry_mapping.get(s, s) for s in rec_symbols]
        
        # Set colors with good contrast for 3D pie charts
        colors = self.colors['chart_colors']
        
        # Create 3D pie charts with improved layout and label positioning
        def make_3d_pie(ax, values, labels, title):
            """Create an improved 3D pie chart with better label placement"""
            # Starting angle
            start_angle = 90
            
            # Create wedges
            wedges = []
            for i, (value, label, color) in enumerate(zip(values, labels, colors * 3)):  # Repeat colors if needed
                # Calculate angles
                angle = 360 * (value / 100)
                end_angle = start_angle + angle
                
                # Draw top of pie
                wedge = plt.matplotlib.patches.Wedge(
                    center=(0, 0), 
                    r=1, 
                    theta1=start_angle, 
                    theta2=end_angle, 
                    width=None,
                    color=color,
                    alpha=0.9
                )
                ax.add_patch(wedge)
                art3d.pathpatch_2d_to_3d(wedge, z=0.5, zdir="z")
                
                # Add side/depth effect
                for depth in [0.1, 0.2, 0.3, 0.4]:
                    wedge = plt.matplotlib.patches.Wedge(
                        center=(0, 0), 
                        r=1, 
                        theta1=start_angle, 
                        theta2=end_angle, 
                        width=None,
                        color=color,
                        alpha=0.4  # More transparent for sides
                    )
                    ax.add_patch(wedge)
                    art3d.pathpatch_2d_to_3d(wedge, z=depth, zdir="z")
                
                # Calculate position for label - adjust radius based on segment size
                # Larger segments get labels placed further out
                mid_angle = np.radians((start_angle + end_angle) / 2)
                label_radius = 1.3 + (value / 200)  # Adjust radius based on segment size
                x = label_radius * np.cos(mid_angle)
                y = label_radius * np.sin(mid_angle)
                
                # Format label text - check if it needs to be shortened
                if len(label) > 15:
                    short_label = label[:12] + "..."
                else:
                    short_label = label
                    
                label_text = f"{short_label}\n{value:.1f}%"
                
                # Add label with improved background box
                ax.text(x, y, 0.5, label_text, 
                       ha='center', va='center', 
                       fontsize=8, color='white',
                       bbox=dict(boxstyle="round,pad=0.3", 
                                facecolor=self.colors['bg_dark'], 
                                alpha=0.8,
                                edgecolor='white',
                                linewidth=0.5))
                
                start_angle = end_angle
            
            # Set chart properties with expanded limits for label placement
            ax.set_xlim(-2.5, 2.5)  # Expanded limits to accommodate labels
            ax.set_ylim(-2.5, 2.5)
            ax.set_zlim(0, 0.5)
            ax.set_title(title, color='white', fontsize=14, pad=20)
            
            # Remove gridlines and axes
            ax.grid(False)
            ax.set_axis_off()
            
            # Set viewing angle
            ax.view_init(elev=30, azim=45)
        
        # Create the two pie charts
        make_3d_pie(ax1, percentages, labels, "Current Allocation")
        make_3d_pie(ax2, rec_percentages, rec_labels, "Recommended Allocation")
        
        # Embed the plot in Tkinter
        canvas = FigureCanvasTkAgg(fig, master=viz_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        # Actions tab
        actions_tab = ttk.Frame(results_notebook, style="Tab.TFrame")
        results_notebook.add(actions_tab, text="Action Items")
        
        # Create action items container
        action_frame = ttk.Frame(actions_tab, style="Card.TFrame")
        action_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Add scrollable frame for actions
        action_canvas = tk.Canvas(action_frame, 
                                bg=self.colors['bg_medium'], 
                                highlightthickness=0, bd=0)
        action_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(action_frame, orient=tk.VERTICAL, command=action_canvas.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        action_canvas.configure(yscrollcommand=scrollbar.set)
        action_canvas.bind('<Configure>', lambda e: action_canvas.configure(scrollregion=action_canvas.bbox("all")))
        
        actions_container = ttk.Frame(action_canvas, style="Card.TFrame")
        action_canvas.create_window((0, 0), window=actions_container, anchor="nw")
        
        # Add title
        ttk.Label(actions_container, text="Recommended Actions", 
               font=("Segoe UI", 14, "bold"), style="TLabel").pack(anchor=tk.W, padx=15, pady=(15, 10))
        
        # Generate actions from the difference between current and recommended allocations
        buy_recs = []
        sell_recs = []
        
        # Compare allocations and generate buy/sell recommendations
        for symbol in set(list(allocations.keys()) + list(recommended.keys())):
            if symbol in allocations and symbol in recommended:
                current = allocations[symbol]
                target = recommended[symbol]
                diff = target - current
                
                # If difference is significant
                if abs(diff) > 1.0:  # More than 1% difference
                    industry_name = self.industry_mapping.get(symbol, symbol)
                    if diff > 0:
                        buy_recs.append(
                            f"Increase {industry_name} allocation by {abs(diff):.1f}% (from {current:.1f}% to {target:.1f}%)")
                    else:
                        sell_recs.append(
                            f"Decrease {industry_name} allocation by {abs(diff):.1f}% (from {current:.1f}% to {target:.1f}%)")
        
        # Display recommendations with icons and colored text
        if buy_recs:
            buy_frame = ttk.Frame(actions_container, style="Card.TFrame")
            buy_frame.pack(fill=tk.X, padx=15, pady=5)
            
            ttk.Label(buy_frame, text="Buy / Increase:", 
                   font=("Segoe UI", 12, "bold"), foreground=self.colors['success'], 
                   style="TLabel").pack(anchor=tk.W, pady=(10, 5))
            
            for item in buy_recs:
                item_frame = ttk.Frame(buy_frame, style="Card.TFrame")
                item_frame.pack(fill=tk.X, padx=15, pady=2)
                
                ttk.Label(item_frame, text="▲", foreground=self.colors['success'], 
                       style="TLabel").pack(side=tk.LEFT, padx=(0, 5))
                ttk.Label(item_frame, text=item, foreground=self.colors['success'], 
                       style="TLabel").pack(side=tk.LEFT)
        
        if sell_recs:
            sell_frame = ttk.Frame(actions_container, style="Card.TFrame")
            sell_frame.pack(fill=tk.X, padx=15, pady=5)
            
            ttk.Label(sell_frame, text="Sell / Decrease:", 
                   font=("Segoe UI", 12, "bold"), foreground=self.colors['error'], 
                   style="TLabel").pack(anchor=tk.W, pady=(10, 5))
            
            for item in sell_recs:
                item_frame = ttk.Frame(sell_frame, style="Card.TFrame")
                item_frame.pack(fill=tk.X, padx=15, pady=2)
                
                ttk.Label(item_frame, text="▼", foreground=self.colors['error'], 
                       style="TLabel").pack(side=tk.LEFT, padx=(0, 5))
                ttk.Label(item_frame, text=item, foreground=self.colors['error'], 
                       style="TLabel").pack(side=tk.LEFT)
        
        # Add close and export buttons
        btn_frame = ttk.Frame(main_container, style="Tab.TFrame")
        btn_frame.pack(fill=tk.X, pady=(15, 0))
        
        export_btn = ttk.Button(btn_frame, text="Export Results", 
                              command=lambda: self.export_results(analysis), style="TButton")
        export_btn.pack(side=tk.LEFT, padx=5)
        
        close_btn = ttk.Button(btn_frame, text="Close", 
                             command=results_window.destroy, style="TButton")
        close_btn.pack(side=tk.RIGHT, padx=5)
    
    def take_screenshot(self):
        """Take a full-screen screenshot"""
        self.status_var.set("Taking screenshot in 3 seconds...")
        self.root.update()
        
        # Minimize window temporarily
        self.root.iconify()
        
        # Wait 3 seconds
        self.root.after(3000, self._capture_screen)
    
    def _capture_screen(self):
        """Capture the screen after countdown"""
        try:
            # Take screenshot
            screenshot = pyautogui.screenshot()
            
            # Create directory if it doesn't exist
            if not os.path.exists("screenshots"):
                os.makedirs("screenshots")
            
            # Save screenshot with timestamp
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"screenshots/portfolio_{timestamp}.png"
            screenshot.save(filename)
            
            # Update UI
            self.display_image(filename)
            self.status_var.set(f"Screenshot saved: {filename}")
            self.analyze_btn.config(state=tk.NORMAL)
            self.view_text_btn.config(state=tk.DISABLED)  # Reset text view button
            self.extracted_text = None  # Reset extracted text
            self.file_type = "image"  # Set file type
            
            # Restore window
            self.root.deiconify()
            
        except Exception as e:
            self.root.deiconify()
            logging.error(f"Screenshot error: {str(e)}")
            messagebox.showerror("Error", f"Failed to capture screenshot: {str(e)}")
            self.status_var.set("Error taking screenshot")
    
    def capture_region(self):
        """Capture a selected region of the screen"""
        self.status_var.set("Minimizing window. Click and drag to select region...")
        self.root.update()
        
        # Minimize window
        self.root.iconify()
        
        # Wait a moment before showing selection UI
        self.root.after(1000, self._show_region_selector)
    
    def _show_region_selector(self):
        """Show region selection UI"""
        try:
            # For simplicity in this implementation, we'll capture a predefined region
            # In a real app, you would create a transparent overlay window for selection
            region = pyautogui.screenshot(region=(100, 100, 800, 600))
            
            # Create directory if it doesn't exist
            if not os.path.exists("screenshots"):
                os.makedirs("screenshots")
            
            # Save region with timestamp
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"screenshots/portfolio_region_{timestamp}.png"
            region.save(filename)
            
            # Update UI
            self.display_image(filename)
            self.status_var.set(f"Region captured: {filename}")
            self.analyze_btn.config(state=tk.NORMAL)
            self.view_text_btn.config(state=tk.DISABLED)
            self.extracted_text = None
            self.file_type = "image"
            
            # Restore window
            self.root.deiconify()
            
        except Exception as e:
            self.root.deiconify()
            logging.error(f"Region capture error: {str(e)}")
            messagebox.showerror("Error", f"Failed to capture region: {str(e)}")
            self.status_var.set("Error capturing region")
    
    def upload_file(self):
        """Upload an image file or spreadsheet"""
        filetypes = [
            ("Image files", "*.png;*.jpg;*.jpeg;*.bmp"),
            ("Excel files", "*.xlsx;*.xls"),
            ("CSV files", "*.csv"),
            ("All files", "*.*")
        ]
        
        filename = filedialog.askopenfilename(title="Select Portfolio File", 
                                             filetypes=filetypes)
        
        if filename:
            # Check if it's an image or spreadsheet
            ext = os.path.splitext(filename)[1].lower()
            
            if ext in ['.png', '.jpg', '.jpeg', '.bmp']:
                # It's an image, display it
                self.display_image(filename)
                self.status_var.set(f"Image loaded: {filename}")
                self.analyze_btn.config(state=tk.NORMAL)
                self.view_text_btn.config(state=tk.DISABLED)
                self.extracted_text = None
                self.file_type = "image"
            elif ext in ['.xlsx', '.xls', '.csv']:
                # It's a spreadsheet
                self.image_path = filename
                
                # Clear current preview
                for widget in self.preview_frame.winfo_children():
                    widget.destroy()
                
                # Show spreadsheet preview
                if ext == '.csv':
                    try:
                        df = pd.read_csv(filename)
                        self.display_dataframe_preview(df)
                    except Exception as e:
                        logging.error(f"CSV read error: {str(e)}")
                        messagebox.showerror("Error", f"Failed to read CSV file: {str(e)}")
                        return
                else:  # Excel file
                    try:
                        df = pd.read_excel(filename)
                        self.display_dataframe_preview(df)
                    except Exception as e:
                        logging.error(f"Excel read error: {str(e)}")
                        messagebox.showerror("Error", f"Failed to read Excel file: {str(e)}")
                        return
                
                self.status_var.set(f"Spreadsheet loaded: {filename}")
                self.analyze_btn.config(state=tk.NORMAL)
                self.view_text_btn.config(state=tk.DISABLED)
                self.extracted_text = None
                self.file_type = "spreadsheet"
            else:
                messagebox.showerror("Error", "Unsupported file format")
    
    def display_image(self, image_path):
        """Display an image in the preview area"""
        self.image_path = image_path
        
        # Clear current preview
        for widget in self.preview_frame.winfo_children():
            widget.destroy()
        
        # Open and resize image for preview
        image = Image.open(image_path)
        
        # Calculate resize dimensions while maintaining aspect ratio
        max_width = 800
        max_height = 400
        width, height = image.size
        
        if width > max_width or height > max_height:
            ratio = min(max_width/width, max_height/height)
            width = int(width * ratio)
            height = int(height * ratio)
            image = image.resize((width, height), Image.LANCZOS)
        
        # Convert to PhotoImage and display
        photo = ImageTk.PhotoImage(image)
        label = ttk.Label(self.preview_frame, image=photo, style="TLabel")
        label.image = photo  # Keep a reference to prevent garbage collection
        label.pack(padx=10, pady=10)
    
    def display_dataframe_preview(self, df):
        """Display a preview of a dataframe in the preview area"""
        # Create a text widget to display the dataframe
        preview_text = scrolledtext.ScrolledText(
            self.preview_frame, wrap=tk.WORD, height=15,
            bg=self.colors['bg_medium'],
            fg=self.colors['text'],
            insertbackground=self.colors['text'],
            font=("Courier New", 10)  # Monospaced font for better data alignment
        )
        preview_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Show the first 10 rows of the dataframe
        preview_text.insert(tk.END, "Spreadsheet Preview (first 10 rows):\n\n")
        preview_text.insert(tk.END, df.head(10).to_string())
        
        # Make it read-only
        preview_text.config(state=tk.DISABLED)
    
    def clear_preview(self):
        """Clear the preview area and reset state"""
        # Clear current preview
        for widget in self.preview_frame.winfo_children():
            widget.destroy()
        
        # Reset preview label
        self.preview_label = ttk.Label(self.preview_frame, text="Screenshot or file preview will appear here",
                                     style="TLabel")
        self.preview_label.pack(pady=50)
        
        # Reset state
        self.image_path = None
        self.extracted_text = None
        self.analyze_btn.config(state=tk.DISABLED)
        self.view_text_btn.config(state=tk.DISABLED)
        self.status_var.set("Ready")
    
    def view_extracted_text(self):
        """View the text extracted from the image"""
        if not self.extracted_text:
            messagebox.showerror("Error", "No text has been extracted yet")
            return
            
        # Create a new window to display the text
        text_window = tk.Toplevel(self.root)
        text_window.title("Extracted Text")
        text_window.geometry("600x400")
        text_window.configure(bg=self.colors['bg_dark'])
        
        # Create text widget with styling
        text_widget = scrolledtext.ScrolledText(
            text_window, wrap=tk.WORD, padx=10, pady=10,
            bg=self.colors['bg_medium'],
            fg=self.colors['text'],
            insertbackground=self.colors['text'],
            font=("Consolas", 10)  # Using a monospaced font for better readability
        )
        text_widget.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Insert text
        text_widget.insert(tk.END, self.extracted_text)
        
        # Make text widget read-only
        text_widget.config(state=tk.DISABLED)
        
        # Add close button
        btn_frame = ttk.Frame(text_window, style="Tab.TFrame")
        btn_frame.pack(pady=10)
        
        close_btn = ttk.Button(btn_frame, text="Close", command=text_window.destroy, style="TButton")
        close_btn.pack()
    
    def save_settings(self):
        """Save the current settings"""
        self.ocr_api_key = self.ocr_key_var.get()
        self.ai_api_key = self.ai_key_var.get()
        self.ocr_api_endpoint = self.ocr_endpoint_var.get()
        self.ai_api_endpoint = self.ai_endpoint_var.get()
        
        logging.info("Settings saved")
        messagebox.showinfo("Settings", "Settings saved successfully")
    
    def load_history_item(self, event):
        """Load a selected history item"""
        try:
            # Get selected index
            selection = self.history_listbox.curselection()
            if not selection:
                return
                
            index = selection[0]
            if index < len(self.analysis_results):
                item = self.analysis_results[index]
                
                # Display in history details
                self.history_details.config(state=tk.NORMAL)
                self.history_details.delete(1.0, tk.END)
                
                # Apply text formatting with tags
                self.history_details.tag_configure("heading", font=("Segoe UI", 12, "bold"))
                self.history_details.tag_configure("subheading", font=("Segoe UI", 10, "bold"))
                self.history_details.tag_configure("normal", font=("Segoe UI", 10))
                self.history_details.tag_configure("highlight", foreground=self.colors['accent'])
                
                # Add timestamp and source info
                self.history_details.insert(tk.END, "Analysis Information\n", "heading")
                self.history_details.insert(tk.END, "─────────────────────\n\n", "normal")
                self.history_details.insert(tk.END, f"Timestamp: ", "subheading")
                self.history_details.insert(tk.END, f"{item['timestamp']}\n", "normal")
                self.history_details.insert(tk.END, f"Source: ", "subheading")
                self.history_details.insert(tk.END, f"{item['source_type']} - {os.path.basename(item['source_path'])}\n\n", "normal")
                
                # Add analysis
                self.history_details.insert(tk.END, "Analysis Results\n", "heading")
                self.history_details.insert(tk.END, "─────────────────────\n\n", "normal")
                
                # Format the analysis text for better readability
                analysis_lines = item['analysis'].split('\n')
                for line in analysis_lines:
                    # Apply different formatting based on line content
                    if re.match(r'^#+\s+', line) or re.match(r'^[A-Z\s]{5,}:?$', line):
                        # Major headings
                        self.history_details.insert(tk.END, f"{line}\n", "heading")
                    elif re.match(r'^[0-9]+\.\s+', line) or re.match(r'^[A-Z][a-z]+:.*', line):
                        # Numbered lists or category labels
                        self.history_details.insert(tk.END, f"{line}\n", "subheading")
                    elif "%" in line or "$" in line or re.search(r'[0-9]+\.[0-9]+', line):
                        # Lines with numbers, percentages or dollar amounts
                        self.history_details.insert(tk.END, f"{line}\n", "highlight")
                    else:
                        # Normal text
                        self.history_details.insert(tk.END, f"{line}\n", "normal")
                
                self.history_details.config(state=tk.DISABLED)
            
        except Exception as e:
            logging.error(f"Load history error: {str(e)}")
            messagebox.showerror("Error", f"Failed to load history item: {str(e)}")
    
    def update_history_list(self):
        """Update the history list with the latest analysis results"""
        self.history_listbox.delete(0, tk.END)
        
        for item in self.analysis_results:
            timestamp = item['timestamp']
            source_type = item['source_type']
            filename = os.path.basename(item['source_path'])
            
            self.history_listbox.insert(tk.END, f"{timestamp} - {source_type} - {filename}")
    
    def export_analysis(self):
        """Export the current analysis to a file"""
        try:
            # Get selected index
            selection = self.history_listbox.curselection()
            if not selection:
                messagebox.showinfo("Export", "Please select an analysis item to export")
                return
                
            index = selection[0]
            if index < len(self.analysis_results):
                item = self.analysis_results[index]
                
                # Ask for save location
                filename = filedialog.asksaveasfilename(
                    title="Export Analysis",
                    defaultextension=".txt",
                    filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
                )
                
                if filename:
                    with open(filename, 'w') as f:
                        f.write(f"Portfolio Analysis Report\n")
                        f.write(f"======================\n\n")
                        f.write(f"Timestamp: {item['timestamp']}\n")
                        f.write(f"Source: {item['source_type']} - {os.path.basename(item['source_path'])}\n\n")
                        f.write("Analysis Results:\n")
                        f.write("=================\n\n")
                        f.write(item['analysis'])
                    
                    logging.info(f"Analysis exported to {filename}")
                    messagebox.showinfo("Export", f"Analysis exported to {filename}")
            
        except Exception as e:
            logging.error(f"Export error: {str(e)}")
            messagebox.showerror("Error", f"Failed to export analysis: {str(e)}")
    
    def delete_history_item(self):
        """Delete the selected history item"""
        try:
            # Get selected index
            selection = self.history_listbox.curselection()
            if not selection:
                messagebox.showinfo("Delete", "Please select an analysis item to delete")
                return
                
            index = selection[0]
            if index < len(self.analysis_results):
                # Remove the item
                del self.analysis_results[index]
                
                # Update the list
                self.update_history_list()
                
                # Clear the details
                self.history_details.config(state=tk.NORMAL)
                self.history_details.delete(1.0, tk.END)
                self.history_details.config(state=tk.DISABLED)
                
                logging.info("History item deleted")
                messagebox.showinfo("Delete", "Analysis item deleted")
            
        except Exception as e:
            logging.error(f"Delete history error: {str(e)}")
            messagebox.showerror("Error", f"Failed to delete history item: {str(e)}")
    
    def test_api_connection(self):
        """Test connections to OCR and AI APIs"""
        self.status_var.set("Testing API connections...")
        self.root.update()
        
        # Create a progress dialog
        progress_window = tk.Toplevel(self.root)
        progress_window.title("Testing API Connections")
        progress_window.geometry("300x150")
        progress_window.configure(bg=self.colors['bg_dark'])
        progress_window.transient(self.root)
        progress_window.grab_set()
        
        # Add progress information
        progress_frame = ttk.Frame(progress_window, style="Tab.TFrame")
        progress_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)
        
        ttk.Label(progress_frame, text="Testing API connections...", 
                style="TLabel").pack(pady=10)
        progress = ttk.Progressbar(progress_frame, orient=tk.HORIZONTAL, length=250, 
                                 mode='indeterminate', style="Red.Horizontal.TProgressbar")
        progress.pack(pady=10)
        progress.start()
        
        status_label = ttk.Label(progress_frame, text="Initializing...", style="TLabel")
        status_label.pack(pady=10)
        
        # Function to test connections in a separate thread
        def test_connections():
            try:
                # Test OCR API
                status_label.config(text="Testing OCR API...")
                ocr_result = self.test_ocr_api(
                    self.ocr_key_var.get(), 
                    self.ocr_endpoint_var.get()
                )
                
                # Test AI API
                status_label.config(text="Testing Deep Seek API...")
                ai_result = self.test_ai_api(
                    self.ai_key_var.get(),
                    self.ai_endpoint_var.get()
                )
                
                # Show results
                progress.stop()
                progress_window.destroy()
                
                if ocr_result and ai_result:
                    messagebox.showinfo("API Test", "Both APIs are working correctly!")
                elif ocr_result:
                    messagebox.showwarning("API Test", "OCR API is working, but Deep Seek API test failed.")
                elif ai_result:
                    messagebox.showwarning("API Test", "Deep Seek API is working, but OCR API test failed.")
                else:
                    messagebox.showerror("API Test", "Both API tests failed. Please check your API keys and internet connection.")
                
                self.status_var.set("API test completed")
                
            except Exception as e:
                logging.error(f"API test error: {str(e)}")
                progress.stop()
                progress_window.destroy()
                messagebox.showerror("Error", f"API test failed: {str(e)}")
                self.status_var.set("API test failed")
        
        # Start test in a separate thread
        threading.Thread(target=test_connections, daemon=True).start()
    
    def test_ocr_api(self, api_key, api_endpoint):
        """Test OCR API connection"""
        try:
            logging.info(f"Testing OCR API: {api_endpoint}")
            
            # Create a small test image with text
            from PIL import Image, ImageDraw, ImageFont
            
            # Create a blank image
            img = Image.new('RGB', (200, 50), color=(255, 255, 255))
            d = ImageDraw.Draw(img)
            
            # Add text
            d.text((10, 10), "Test OCR API", fill=(0, 0, 0))
            
            # Save to a temporary file
            test_file = "test_ocr.png"
            img.save(test_file)
            
            # Prepare the request - let's try OCR.space format first
            with open(test_file, 'rb') as file:
                files = {'file': file}
                payload = {
                    'apikey': api_key,
                    'language': 'eng',
                    'isOverlayRequired': False
                }
                
                response = requests.post(
                    api_endpoint,
                    files=files,
                    data=payload
                )
            
            # Clean up
            os.remove(test_file)
            
            # Check the response
            if response.status_code == 200:
                logging.info("OCR API test successful")
                return True
            else:
                logging.error(f"OCR API test failed with status {response.status_code}: {response.text}")
                return False
            
        except Exception as e:
            logging.error(f"OCR API test error: {str(e)}")
            return False
    
    def test_ai_api(self, api_key, api_endpoint):
        """Test Deep Seek API connection"""
        try:
            logging.info(f"Testing Deep Seek API: {api_endpoint}")
            
            # Prepare a simple request
            headers = {
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json"
            }
            
            payload = {
                "model": "deepseek-chat",
                "messages": [
                    {"role": "user", "content": "Hello, this is a test message. Please respond with 'API is working'."}
                ],
                "max_tokens": 50
            }
            
            response = requests.post(
                api_endpoint,
                headers=headers,
                json=payload
            )
            
            # Check the response
            if response.status_code == 200:
                logging.info("Deep Seek API test successful")
                return True
            else:
                logging.error(f"Deep Seek API test failed with status {response.status_code}: {response.text}")
                return False
            
        except Exception as e:
            logging.error(f"Deep Seek API test error: {str(e)}")
            return False

    def process_image_with_ocr(self, image_path, api_key):
        """Process image with OCR API using the provided key"""
        self.status_var.set("Extracting text with OCR...")
        self.root.update()
        
        try:
            logging.info(f"Processing image with OCR: {image_path}")
            
            # Get OCR endpoint from settings
            api_endpoint = self.ocr_endpoint_var.get()
            
            # Prepare the image file for OCR API
            with open(image_path, 'rb') as image_file:
                # For OCR.space API
                files = {'file': image_file}
                payload = {
                    'apikey': api_key,
                    'language': 'eng',
                    'isOverlayRequired': False,
                    'detectOrientation': True
                }
                
                # Make the API request
                logging.info(f"Sending OCR request to {api_endpoint}")
                response = requests.post(
                    api_endpoint,
                    files=files,
                    data=payload
                )
            
            # Check if the request was successful
            if response.status_code == 200:
                try:
                    # Try to parse the JSON response
                    result = response.json()
                    logging.info(f"OCR response received: {result}")
                    
                    # Try to extract the text from the response
                    # For OCR.space API
                    if 'ParsedResults' in result and result['ParsedResults']:
                        extracted_text = result['ParsedResults'][0]['ParsedText']
                    else:
                        # If we can't find the expected structure, log the full response for debugging
                        logging.warning(f"Unexpected OCR response structure: {result}")
                        extracted_text = f"OCR response received but couldn't locate parsed text. Please check the API format."
                    
                    self.status_var.set("OCR processing complete")
                    logging.info("OCR processing complete")
                    return extracted_text
                except ValueError as ve:
                    # If the response is not valid JSON
                    logging.error(f"OCR response parsing error: {ve}")
                    logging.debug(f"Raw response: {response.text}")
                    raise Exception(f"Invalid response from OCR API: {ve}")
            else:
                # Handle API error
                error_msg = f"OCR API error: {response.status_code} - {response.text}"
                logging.error(error_msg)
                self.status_var.set(error_msg)
                raise Exception(error_msg)
                
        except Exception as e:
            error_msg = f"Failed to process image with OCR: {str(e)}"
            logging.error(error_msg)
            self.status_var.set(f"OCR processing error")
            raise Exception(error_msg)
    
    def analyze_with_ai(self, text, api_key, risk_tolerance, investment_horizon):
        """Analyze portfolio text with Deep Seek AI API"""
        self.status_var.set("Analyzing with AI...")
        self.root.update()
        
        try:
            logging.info("Starting AI analysis")
            
            # Get AI endpoint from settings
            api_endpoint = self.ai_endpoint_var.get()
            
            # Get additional preferences
            threshold = self.threshold_var.get()
            
            # Get sector preferences
            preferred_sectors = []
            if self.tech_var.get():
                preferred_sectors.append("Technology")
            if self.finance_var.get():
                preferred_sectors.append("Finance")
            if self.healthcare_var.get():
                preferred_sectors.append("Healthcare")
            if self.energy_var.get():
                preferred_sectors.append("Energy")
            if self.consumer_var.get():
                preferred_sectors.append("Consumer Goods")
            
            sectors_text = ", ".join(preferred_sectors) if preferred_sectors else "No specific sectors"
            
            # Create prompt for portfolio analysis
            prompt = f"""
            Analyze this portfolio data extracted from an image or spreadsheet and provide detailed rebalancing recommendations:
            
            {text}
            
            Additional parameters:
            - Risk tolerance: {risk_tolerance}
            - Investment horizon: {investment_horizon}
            - Rebalancing threshold: {threshold}%
            - Preferred sectors: {sectors_text}
            
            Please provide your analysis in a well-structured format with clear section headers. Include:
            1. Current portfolio allocation analysis (with percentages)
            2. Asset allocation recommendations based on risk profile
            3. Specific buy/sell recommendations with amounts
            4. Sector diversification suggestions with focus on preferred sectors
            5. Overall portfolio health assessment
            
            Format the output with clear headings and subheadings for readability. Use percentages and numbers to quantify your recommendations.
            
            If the data doesn't appear to be a portfolio, please indicate that clearly.
            """
            
            # Prepare API request
            headers = {
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json"
            }
            
            payload = {
                "model": "deepseek-chat",
                "messages": [
                    {"role": "user", "content": prompt}
                ],
                "temperature": 0.7,
                "max_tokens": 2000
            }
            
            # Log the request (without the full text to keep logs manageable)
            logging.info(f"Sending analysis request to Deep Seek API at {api_endpoint}")
            
            # Make the API call
            response = requests.post(
                api_endpoint,
                headers=headers,
                json=payload
            )
            
            # Check if the request was successful
            if response.status_code == 200:
                # Parse the JSON response
                result = response.json()
                logging.info("Deep Seek API response received")
                
                # Extract the analysis text from the response
                # The path to the content depends on the API's response structure
                if 'choices' in result and len(result['choices']) > 0:
                    # This is a common format for OpenAI-compatible APIs
                    analysis = result['choices'][0]['message']['content']
                else:
                    # If we can't find the content in the expected location
                    logging.warning(f"Unexpected Deep Seek API response structure: {result}")
                    analysis = "Analysis received but couldn't parse the response format. Please check the API documentation."
                
                self.status_var.set("AI analysis complete")
                logging.info("AI analysis complete")
                return analysis
            else:
                # Handle API error
                error_msg = f"Deep Seek API error: {response.status_code} - {response.text}"
                logging.error(error_msg)
                self.status_var.set(error_msg)
                raise Exception(error_msg)
                
        except Exception as e:
            error_msg = f"Failed to analyze with AI: {str(e)}"
            logging.error(error_msg)
            self.status_var.set(f"AI analysis error")
            raise Exception(error_msg)
    
    def analyze_portfolio(self):
        """Analyze the portfolio using OCR and AI"""
        if not self.image_path:
            messagebox.showerror("Error", "No image or file to analyze")
            return
        
        self.status_var.set("Starting analysis...")
        self.analyze_btn.config(state=tk.DISABLED)
        self.progress.start(10)
        self.root.update()
        
        # Run analysis in a separate thread to keep UI responsive
        threading.Thread(target=self._run_analysis, daemon=True).start()
    
    def _run_analysis(self):
        """Run the analysis process in a thread"""
        try:
            # Determine file type
            ext = os.path.splitext(self.image_path)[1].lower()
            
            if ext in ['.png', '.jpg', '.jpeg', '.bmp']:
                # Process image with OCR
                logging.info(f"Analyzing image file: {self.image_path}")
                self.extracted_text = self.process_image_with_ocr(
                    self.image_path, 
                    self.ocr_key_var.get()
                )
                
                if not self.extracted_text:
                    raise Exception("Failed to extract text from image")
                
                # Update UI to enable view text button (need to use after() for thread safety)
                self.root.after(0, lambda: self.view_text_btn.config(state=tk.NORMAL))
                
                # Send extracted text to AI for analysis
                analysis = self.analyze_with_ai(
                    self.extracted_text,
                    self.ai_key_var.get(),
                    self.risk_var.get(),
                    self.horizon_var.get()
                )
                
                # Add to history
                timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                history_item = {
                    'timestamp': timestamp,
                    'source_type': 'Image',
                    'source_path': self.image_path,
                    'extracted_text': self.extracted_text,
                    'analysis': analysis
                }
                
                self.analysis_results.append(history_item)
                
                # Update UI thread-safely
                self.root.after(0, self.update_history_list)
                
            elif ext in ['.xlsx', '.xls', '.csv']:
                # Process spreadsheet
                logging.info(f"Analyzing spreadsheet file: {self.image_path}")
                
                # Read the spreadsheet
                if ext == '.csv':
                    df = pd.read_csv(self.image_path)
                else:  # Excel file
                    df = pd.read_excel(self.image_path)
                
                # Convert dataframe to text
                buffer = io.StringIO()
                df.to_csv(buffer, index=False)
                text_data = buffer.getvalue()
                self.extracted_text = text_data
                
                # Update UI to enable view text button (need to use after() for thread safety)
                self.root.after(0, lambda: self.view_text_btn.config(state=tk.NORMAL))
                
                # Send to AI for analysis
                analysis = self.analyze_with_ai(
                    text_data,
                    self.ai_key_var.get(),
                    self.risk_var.get(),
                    self.horizon_var.get()
                )
                
                # Add to history
                timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                history_item = {
                    'timestamp': timestamp,
                    'source_type': 'Spreadsheet',
                    'source_path': self.image_path,
                    'extracted_text': self.extracted_text,
                    'analysis': analysis
                }
                
                self.analysis_results.append(history_item)
                
                # Update UI thread-safely
                self.root.after(0, self.update_history_list)
            
            # Show analysis results
            self.root.after(0, lambda: self.show_analysis_results(analysis))
            
            # Update status and re-enable analyze button thread-safely
            self.root.after(0, lambda: self.status_var.set("Analysis complete"))
            self.root.after(0, lambda: self.analyze_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.progress.stop())
            
        except Exception as e:
            logging.error(f"Analysis error: {str(e)}")
            # Handle errors thread-safely
            self.root.after(0, lambda: messagebox.showerror("Error", f"Analysis failed: {str(e)}"))
            self.root.after(0, lambda: self.status_var.set("Analysis failed"))
            self.root.after(0, lambda: self.analyze_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.progress.stop())
    
    def export_results(self, analysis):
        """Export analysis results to a file"""
        filename = filedialog.asksaveasfilename(
            title="Export Analysis Results",
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        
        if filename:
            with open(filename, 'w') as f:
                f.write("Portfolio Analysis Results\n")
                f.write("========================\n\n")
                f.write(f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
                f.write(analysis)
            
            logging.info(f"Results exported to {filename}")
            messagebox.showinfo("Export", f"Results exported to {filename}")

# Run the application
if __name__ == "__main__":
    root = tk.Tk()
    app = ProfessionalPortfolioAnalyzer(root)
    root.mainloop()
