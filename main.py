import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import time
from datetime import datetime, timedelta
import json
import os
import pandas as pd
import sys

def get_resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    
    return os.path.join(base_path, relative_path)

class GamingLoungeManager:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Gaming Lounge Manager")
        self.root.attributes('-zoomed', True)  # Cross-platform full screen
        
        # Configuration
        self.config = {
            "playstation_rate": 6000.0,  # per hour
            "services": {
                "coffee": 2500.0,
                "matte": 5000.0,
                "tea": 2000.0,
                "shisha": 5000.0
            },
            "offers": {
                "enabled": True,
                "2_hour_rate": 5000.0,  # rate when 2+ hours
                "3_hour_rate": 4666.0   # rate when 3+ hours
            }
        }
        
        # Update config file path
        self.config_file = get_resource_path("config.json")
        
        # Load config from file if exists
        self.load_config()
        
        # PlayStation sessions
        self.sessions = {
            "PS1": {"active": False, "start_time": None, "customer_name": "", "services": []},
            "PS2": {"active": False, "start_time": None, "customer_name": "", "services": []},
            "PS3": {"active": False, "start_time": None, "customer_name": "", "services": []},
            "PS4": {"active": False, "start_time": None, "customer_name": "", "services": []}
        }
        
        # Database file
        self.db_file = "gaming_lounge_db.xlsx"
        self.init_database()
        
        self.setup_ui()
        self.update_timer()
        
    def init_database(self):
        """Initialize Excel database if it doesn't exist"""
        if not os.path.exists(self.db_file):
            df = pd.DataFrame(columns=[
                'Date', 'Time', 'PlayStation', 'Customer', 'Duration_Hours', 
                'PS_Cost', 'Services', 'Service_Cost', 'Total_Cost'
            ])
            df.to_excel(self.db_file, index=False)
        
    def setup_ui(self):
        """Setup the main UI"""
        # Create notebook for tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Main tab
        self.main_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.main_tab, text="PlayStation Manager")
        
        # Database tab
        self.database_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.database_tab, text="Database")
        
        # Settings tab
        self.settings_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.settings_tab, text="Settings")
        
        # Setup each tab
        self.setup_main_tab()
        self.setup_database_tab()
        self.setup_settings_tab()

    def setup_main_tab(self):
        """Setup the main PlayStation management tab"""
        # Main frame
        main_frame = ttk.Frame(self.main_tab, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_label = ttk.Label(main_frame, text="Gaming Lounge Manager", 
                               font=("Arial", 20, "bold"))
        title_label.grid(row=0, column=0, columnspan=4, pady=(0, 20))
        
        # Configure grid weights
        for i in range(4):
            main_frame.columnconfigure(i, weight=1)
        
        # PlayStation controls
        self.ps_frames = {}
        for i, ps_name in enumerate(["PS1", "PS2", "PS3", "PS4"]):
            frame = ttk.LabelFrame(main_frame, text=ps_name, padding="10")
            frame.grid(row=1, column=i, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))
            
            # Status label
            status_label = ttk.Label(frame, text="Available", foreground="green")
            status_label.grid(row=0, column=0, pady=5, columnspan=2)
            
            # Timer label
            timer_label = ttk.Label(frame, text="00:00:00", font=("Arial", 12, "bold"))
            timer_label.grid(row=1, column=0, pady=5, columnspan=2)
            
            # Current cost label
            cost_label = ttk.Label(frame, text="Cost: $0", font=("Arial", 10, "bold"), foreground="blue")
            cost_label.grid(row=2, column=0, pady=5, columnspan=2)
            
            # Apply button
            apply_btn = ttk.Button(frame, text="Apply", 
                                 command=lambda ps=ps_name: self.start_session(ps))
            apply_btn.grid(row=3, column=0, pady=5, sticky=(tk.W, tk.E))
            
            # Done button
            done_btn = ttk.Button(frame, text="Done", 
                                command=lambda ps=ps_name: self.end_session(ps))
            done_btn.grid(row=3, column=1, pady=5, sticky=(tk.W, tk.E))
            
            # Services section for this PS
            services_label = ttk.Label(frame, text="Services:", font=("Arial", 10, "bold"))
            services_label.grid(row=4, column=0, columnspan=2, pady=(10, 5), sticky=tk.W)
            
            # Service buttons
            service_buttons = {}
            for j, (service, price) in enumerate(self.config["services"].items()):
                btn = ttk.Button(frame, text=f"{service.title()}\n${price}", 
                               command=lambda ps=ps_name, srv=service: self.add_service(ps, srv))
                btn.grid(row=5 + j//2, column=j%2, pady=2, padx=2, sticky=(tk.W, tk.E))
                service_buttons[service] = btn
            
            # Services list
            services_listbox = tk.Listbox(frame, height=3, font=("Arial", 8))
            services_listbox.grid(row=7, column=0, columnspan=2, pady=5, sticky=(tk.W, tk.E))
            
            # Remove service button
            remove_service_btn = ttk.Button(frame, text="Remove Selected", 
                                          command=lambda ps=ps_name: self.remove_service(ps))
            remove_service_btn.grid(row=8, column=0, columnspan=2, pady=2, sticky=(tk.W, tk.E))
            
            self.ps_frames[ps_name] = {
                "frame": frame,
                "status_label": status_label,
                "timer_label": timer_label,
                "cost_label": cost_label,
                "apply_btn": apply_btn,
                "done_btn": done_btn,
                "service_buttons": service_buttons,
                "services_listbox": services_listbox,
                "remove_service_btn": remove_service_btn
            }

        # Services Only section
        services_only_frame = ttk.LabelFrame(main_frame, text="Services Only", padding="10")
        services_only_frame.grid(row=1, column=4, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Services Only label
        services_label = ttk.Label(services_only_frame, text="No PlayStation", 
                                  foreground="blue", font=("Arial", 10, "bold"))
        services_label.grid(row=0, column=0, columnspan=2, pady=5)
        
        # Customer name entry
        ttk.Label(services_only_frame, text="Customer:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.services_customer_var = tk.StringVar()
        customer_entry = ttk.Entry(services_only_frame, textvariable=self.services_customer_var, width=15)
        customer_entry.grid(row=1, column=1, pady=2, sticky=(tk.W, tk.E))
        
        # Services section
        services_label = ttk.Label(services_only_frame, text="Services:", font=("Arial", 10, "bold"))
        services_label.grid(row=2, column=0, columnspan=2, pady=(10, 5), sticky=tk.W)
        
        # Service buttons for services only
        self.services_only_buttons = {}
        for j, (service, price) in enumerate(self.config["services"].items()):
            btn = ttk.Button(services_only_frame, text=f"{service.title()}\n${price}", 
                           command=lambda srv=service: self.add_service_only(srv))
            btn.grid(row=3 + j//2, column=j%2, pady=2, padx=2, sticky=(tk.W, tk.E))
            self.services_only_buttons[service] = btn
        
        # Add to pending button
        add_pending_btn = ttk.Button(services_only_frame, text="Add to Pending Orders", 
                                   command=self.add_to_pending_orders)
        add_pending_btn.grid(row=6, column=0, columnspan=2, pady=10, sticky=(tk.W, tk.E))
        
        # Current order display
        current_order_label = ttk.Label(services_only_frame, text="Current Order:", font=("Arial", 9, "bold"))
        current_order_label.grid(row=7, column=0, columnspan=2, pady=(10, 5), sticky=tk.W)
        
        self.current_order_listbox = tk.Listbox(services_only_frame, height=3, font=("Arial", 8))
        self.current_order_listbox.grid(row=8, column=0, columnspan=2, pady=5, sticky=(tk.W, tk.E))
        
        # Remove service button
        remove_service_btn = ttk.Button(services_only_frame, text="Remove Selected", 
                                      command=self.remove_service_only)
        remove_service_btn.grid(row=9, column=0, columnspan=2, pady=2, sticky=(tk.W, tk.E))
        
        # Clear current order button
        clear_btn = ttk.Button(services_only_frame, text="Clear Current Order", 
                             command=self.clear_current_order)
        clear_btn.grid(row=10, column=0, columnspan=2, pady=2, sticky=(tk.W, tk.E))
        
        # Pending Orders section
        pending_frame = ttk.LabelFrame(main_frame, text="Pending Orders", padding="10")
        pending_frame.grid(row=2, column=0, columnspan=5, padx=5, pady=10, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Pending orders treeview
        pending_columns = ['Customer', 'Services', 'Total', 'Time Added']
        self.pending_tree = ttk.Treeview(pending_frame, columns=pending_columns, show='headings', height=8)
        
        for col in pending_columns:
            self.pending_tree.heading(col, text=col)
            if col == 'Total':
                self.pending_tree.column(col, width=80, anchor='center')
            elif col == 'Time Added':
                self.pending_tree.column(col, width=100, anchor='center')
            else:
                self.pending_tree.column(col, width=200)
        
        self.pending_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Pending orders scrollbar
        pending_scrollbar = ttk.Scrollbar(pending_frame, orient=tk.VERTICAL, command=self.pending_tree.yview)
        self.pending_tree.configure(yscrollcommand=pending_scrollbar.set)
        pending_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Pending orders buttons
        pending_buttons_frame = ttk.Frame(pending_frame)
        pending_buttons_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Button(pending_buttons_frame, text="Generate Bill for Selected", 
                  command=self.generate_bill_for_pending).pack(side=tk.LEFT, padx=5)
        ttk.Button(pending_buttons_frame, text="Remove Selected Order", 
                  command=self.remove_pending_order).pack(side=tk.LEFT, padx=5)
        ttk.Button(pending_buttons_frame, text="Add More Services", 
                  command=self.add_more_services_to_pending).pack(side=tk.LEFT, padx=5)
        
        # Initialize data structures
        self.services_only_list = []
        self.pending_orders = []

    def setup_settings_tab(self):
        # Settings frame
        settings_frame = ttk.Frame(self.settings_tab, padding="20")
        settings_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_label = ttk.Label(settings_frame, text="Price Configuration", 
                               font=("Arial", 16, "bold"))
        title_label.pack(pady=(0, 20))
        
        # PlayStation rate section
        ps_frame = ttk.LabelFrame(settings_frame, text="PlayStation Rate", padding="10")
        ps_frame.pack(fill=tk.X, pady=(0, 20))
        
        ttk.Label(ps_frame, text="Rate per hour ($):").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.ps_rate_var = tk.StringVar(value=str(self.config["playstation_rate"]))
        ps_rate_entry = ttk.Entry(ps_frame, textvariable=self.ps_rate_var, width=10)
        ps_rate_entry.grid(row=0, column=1, sticky=tk.W)
        
        # Services section
        services_frame = ttk.LabelFrame(settings_frame, text="Service Prices", padding="10")
        services_frame.pack(fill=tk.X, pady=(0, 20))
        
        self.service_vars = {}
        for i, (service, price) in enumerate(self.config["services"].items()):
            ttk.Label(services_frame, text=f"{service.title()} ($):").grid(row=i, column=0, sticky=tk.W, padx=(0, 10), pady=2)
            var = tk.StringVar(value=str(price))
            entry = ttk.Entry(services_frame, textvariable=var, width=10)
            entry.grid(row=i, column=1, sticky=tk.W, pady=2)
            self.service_vars[service] = var
        
        # Offers section
        offers_frame = ttk.LabelFrame(settings_frame, text="Hour Offers", padding="10")
        offers_frame.pack(fill=tk.X, pady=(0, 20))
        
        # Enable offers checkbox
        self.offers_enabled_var = tk.BooleanVar(value=self.config["offers"]["enabled"])
        offers_checkbox = ttk.Checkbutton(offers_frame, text="Enable Hour Offers", 
                                         variable=self.offers_enabled_var)
        offers_checkbox.grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 10))
        
        # 2 hour offer
        ttk.Label(offers_frame, text="2+ Hour Rate ($):").grid(row=1, column=0, sticky=tk.W, padx=(0, 10), pady=2)
        self.offer_2h_var = tk.StringVar(value=str(self.config["offers"]["2_hour_rate"]))
        offer_2h_entry = ttk.Entry(offers_frame, textvariable=self.offer_2h_var, width=10)
        offer_2h_entry.grid(row=1, column=1, sticky=tk.W, pady=2)
        
        # 3 hour offer
        ttk.Label(offers_frame, text="3+ Hour Rate ($):").grid(row=2, column=0, sticky=tk.W, padx=(0, 10), pady=2)
        self.offer_3h_var = tk.StringVar(value=str(self.config["offers"]["3_hour_rate"]))
        offer_3h_entry = ttk.Entry(offers_frame, textvariable=self.offer_3h_var, width=10)
        offer_3h_entry.grid(row=2, column=1, sticky=tk.W, pady=2)
        
        # Buttons
        button_frame = ttk.Frame(settings_frame)
        button_frame.pack(pady=20)
        
        ttk.Button(button_frame, text="Save Settings", 
                  command=self.save_settings).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(button_frame, text="Reset to Default", 
                  command=self.reset_settings).pack(side=tk.LEFT, padx=5)

    def save_settings(self):
        """Save settings and update UI"""
        # Confirmation dialog
        result = messagebox.askyesno("Confirm Save", "Save these price settings?")
        if not result:
            return
        
        try:
            # Update PlayStation rate
            new_ps_rate = float(self.ps_rate_var.get())
            self.config["playstation_rate"] = new_ps_rate
            
            # Update service prices
            for service, var in self.service_vars.items():
                new_price = float(var.get())
                self.config["services"][service] = new_price
            
            # Update offers
            self.config["offers"]["enabled"] = self.offers_enabled_var.get()
            self.config["offers"]["2_hour_rate"] = float(self.offer_2h_var.get())
            self.config["offers"]["3_hour_rate"] = float(self.offer_3h_var.get())
            
            # Save to config file
            with open(self.config_file, "w") as f:
                json.dump(self.config, f, indent=4)
            
            # Update service buttons in main tab
            self.update_service_buttons()
            
            messagebox.showinfo("Success", "Settings saved successfully!")
            
        except ValueError:
            messagebox.showerror("Error", "Please enter valid numbers for all prices")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save settings: {str(e)}")

    def reset_settings(self):
        """Reset to default settings"""
        default_config = {
            "playstation_rate": 6000.0,
            "services": {
                "coffee": 2500.0,
                "matte": 5000.0,
                "tea": 2000.0,
                "shisha": 5000.0
            },
            "offers": {
                "enabled": True,
                "2_hour_rate": 5000.0,
                "3_hour_rate": 4666.0
            }
        }
        
        result = messagebox.askyesno("Confirm Reset", "Reset all prices to default values?")
        if result:
            self.config = default_config.copy()
            
            # Update UI
            self.ps_rate_var.set(str(self.config["playstation_rate"]))
            for service, var in self.service_vars.items():
                var.set(str(self.config["services"][service]))
            
            # Update offers UI
            self.offers_enabled_var.set(self.config["offers"]["enabled"])
            self.offer_2h_var.set(str(self.config["offers"]["2_hour_rate"]))
            self.offer_3h_var.set(str(self.config["offers"]["3_hour_rate"]))
            
            self.update_service_buttons()
            messagebox.showinfo("Success", "Settings reset to default values!")

    def load_config_file(self):
        """Load settings from config.json file"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, "r") as f:
                    loaded_config = json.load(f)
                
                # Update config
                if "playstation_rate" in loaded_config:
                    self.config["playstation_rate"] = loaded_config["playstation_rate"]
                    self.ps_rate_var.set(str(self.config["playstation_rate"]))
                
                if "services" in loaded_config:
                    for service in self.config["services"]:
                        if service in loaded_config["services"]:
                            self.config["services"][service] = loaded_config["services"][service]
                            self.service_vars[service].set(str(self.config["services"][service]))
                
                # Update offers
                if "offers" in loaded_config:
                    if "enabled" in loaded_config["offers"]:
                        self.config["offers"]["enabled"] = loaded_config["offers"]["enabled"]
                        self.offers_enabled_var.set(self.config["offers"]["enabled"])
                    if "2_hour_rate" in loaded_config["offers"]:
                        self.config["offers"]["2_hour_rate"] = loaded_config["offers"]["2_hour_rate"]
                        self.offer_2h_var.set(str(self.config["offers"]["2_hour_rate"]))
                    if "3_hour_rate" in loaded_config["offers"]:
                        self.config["offers"]["3_hour_rate"] = loaded_config["offers"]["3_hour_rate"]
                        self.offer_3h_var.set(str(self.config["offers"]["3_hour_rate"]))
                
                self.update_service_buttons()
                messagebox.showinfo("Success", "Settings loaded from config.json!")
            else:
                messagebox.showwarning("Warning", "config.json file not found!")
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load config file: {str(e)}")

    def update_service_buttons(self):
        """Update service button texts with new prices"""
        for ps_name, ps_frame in self.ps_frames.items():
            for service, button in ps_frame["service_buttons"].items():
                price = self.config["services"][service]
                button.config(text=f"{service.title()}\n${price}")
        
        # Update services-only buttons
        if hasattr(self, 'services_only_buttons'):
            for service, button in self.services_only_buttons.items():
                price = self.config["services"][service]
                button.config(text=f"{service.title()}\n${price}")

    def load_config(self):
        """Load configuration from file"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, "r") as f:
                    loaded_config = json.load(f)
                # Update config with loaded values
                self.config.update(loaded_config)
        except Exception as e:
            print(f"Could not load config: {e}")
            # Use default config

    def add_service(self, ps_name, service):
        """Add service to PlayStation with confirmation"""
        if not self.sessions[ps_name]["active"]:
            messagebox.showwarning("Warning", f"{ps_name} is not active. Start a session first.")
            return
            
        # Show confirmation dialog
        result = messagebox.askyesno("Confirm Service", 
                                   f"Add {service.title()} (${self.config['services'][service]}) to {ps_name}?")
        
        if result:
            service_entry = {
                "name": service,
                "price": self.config["services"][service],
                "time": datetime.now().strftime("%H:%M:%S")
            }
            
            self.sessions[ps_name]["services"].append(service_entry)
            self.update_services_display(ps_name)
            messagebox.showinfo("Success", f"{service.title()} added to {ps_name}")
    
    def remove_service(self, ps_name):
        """Remove selected service from PlayStation"""
        listbox = self.ps_frames[ps_name]["services_listbox"]
        selection = listbox.curselection()
        
        if not selection:
            messagebox.showwarning("Warning", "Please select a service to remove")
            return
            
        index = selection[0]
        service_name = self.sessions[ps_name]["services"][index]["name"]
        
        result = messagebox.askyesno("Confirm Removal", 
                                   f"Remove {service_name.title()} from {ps_name}?")
        
        if result:
            del self.sessions[ps_name]["services"][index]
            self.update_services_display(ps_name)
    
    def update_services_display(self, ps_name):
        """Update the services listbox display"""
        listbox = self.ps_frames[ps_name]["services_listbox"]
        listbox.delete(0, tk.END)
        
        for service in self.sessions[ps_name]["services"]:
            display_text = f"{service['name'].title()} - ${service['price']} ({service['time']})"
            listbox.insert(tk.END, display_text)

    def start_session(self, ps_name):
        # Confirmation dialog
        result = messagebox.askyesno("Confirm Start", f"Start gaming session on {ps_name}?")
        if not result:
            return
        
        if self.sessions[ps_name]["active"]:
            messagebox.showwarning("Warning", f"{ps_name} is already in use")
            return
        
        self.sessions[ps_name] = {
            "active": True,
            "start_time": time.time(),
            "customer_name": "",
            "services": []
        }
        
        self.ps_frames[ps_name]["status_label"].config(text="In Use", foreground="red")
        self.ps_frames[ps_name]["apply_btn"].config(state="disabled")
        
        messagebox.showinfo("Success", f"{ps_name} started")

    def end_session(self, ps_name):
        # Confirmation dialog
        result = messagebox.askyesno("Confirm End", f"End gaming session on {ps_name}?")
        if not result:
            return
        
        if not self.sessions[ps_name]["active"]:
            messagebox.showwarning("Warning", f"{ps_name} is not in use")
            return
        
        # Calculate duration and cost with offers
        duration = time.time() - self.sessions[ps_name]["start_time"]
        hours = duration / 3600
        ps_cost = self.calculate_ps_cost(hours)
        
        # Calculate services cost
        service_cost = sum(service["price"] for service in self.sessions[ps_name]["services"])
        total_cost = ps_cost + service_cost
        
        # Show bill
        self.show_bill(ps_name, duration, ps_cost, service_cost, total_cost)
    
    def show_bill(self, ps_name, duration, ps_cost, service_cost, total_cost):
        bill_window = tk.Toplevel(self.root)
        bill_window.title("Bill")
        bill_window.geometry("400x700")
        
        # Bill content
        bill_frame = ttk.Frame(bill_window, padding="20")
        bill_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(bill_frame, text="GAMING LOUNGE BILL", 
                 font=("Arial", 16, "bold")).pack(pady=(0, 20))
        
        ttk.Label(bill_frame, text=f"PlayStation: {ps_name}").pack(anchor=tk.W)
        ttk.Label(bill_frame, text=f"Date: {datetime.now().strftime('%Y-%m-%d %H:%M')}").pack(anchor=tk.W)
        
        ttk.Separator(bill_frame, orient='horizontal').pack(fill=tk.X, pady=10)
        
        # Duration in HH:MM format
        hours = int(duration // 3600)
        minutes = int((duration % 3600) // 60)
        duration_str = f"{hours:02d}:{minutes:02d}"
        duration_hours = duration / 3600
        
        ttk.Label(bill_frame, text=f"Duration: {duration_str}").pack(anchor=tk.W)
        ttk.Label(bill_frame, text=f"Base Rate: ${int(self.config['playstation_rate']):,}/hour").pack(anchor=tk.W)
        
        # Calculate normal cost (without offers)
        normal_ps_cost = duration_hours * self.config["playstation_rate"]
        ttk.Label(bill_frame, text=f"Normal PlayStation Cost: ${int(normal_ps_cost):,}").pack(anchor=tk.W)
        
        # Check for offer eligibility
        offer_applied = False
        offer_savings = 0
        
        if duration_hours >= 3:
            offer_rate = self.config["offers"]["3_hour_rate"]
            offer_ps_cost = duration_hours * offer_rate
            offer_savings = normal_ps_cost - offer_ps_cost
            offer_text = "3+ Hour Offer"
            offer_applied = True
        elif duration_hours >= 2:
            offer_rate = self.config["offers"]["2_hour_rate"]
            offer_ps_cost = duration_hours * offer_rate
            offer_savings = normal_ps_cost - offer_ps_cost
            offer_text = "2+ Hour Offer"
            offer_applied = True
        
        # Show offer section if eligible
        if offer_applied:
            ttk.Separator(bill_frame, orient='horizontal').pack(fill=tk.X, pady=10)
            
            # Offer eligibility frame
            offer_frame = ttk.LabelFrame(bill_frame, text=f"🎉 {offer_text} Available!", padding="10")
            offer_frame.pack(fill=tk.X, pady=5)
            
            ttk.Label(offer_frame, text=f"Offer Rate: ${int(offer_rate):,}/hour", 
                     font=("Arial", 10, "bold"), foreground="green").pack(anchor=tk.W)
            ttk.Label(offer_frame, text=f"With Offer: ${int(offer_ps_cost):,}", 
                     font=("Arial", 10, "bold"), foreground="green").pack(anchor=tk.W)
            ttk.Label(offer_frame, text=f"You Save: ${int(offer_savings):,}", 
                     font=("Arial", 10, "bold"), foreground="red").pack(anchor=tk.W)
            
            # Offer selection
            self.apply_offer_var = tk.BooleanVar(value=True)
            offer_checkbox = ttk.Checkbutton(offer_frame, text=f"Apply {offer_text}", 
                                            variable=self.apply_offer_var,
                                            command=lambda: self.update_bill_total(bill_frame, duration_hours, service_cost))
            offer_checkbox.pack(anchor=tk.W, pady=5)
            
            # Store offer details for calculation
            self.current_offer = {
                "normal_cost": normal_ps_cost,
                "offer_cost": offer_ps_cost,
                "savings": offer_savings
            }
        else:
            self.apply_offer_var = tk.BooleanVar(value=False)
            self.current_offer = {"normal_cost": normal_ps_cost, "offer_cost": normal_ps_cost, "savings": 0}
        
        ttk.Separator(bill_frame, orient='horizontal').pack(fill=tk.X, pady=10)
        
        # Services
        ttk.Label(bill_frame, text="Services Used:", font=("Arial", 12, "bold")).pack(anchor=tk.W)
        
        if self.sessions[ps_name]["services"]:
            for service in self.sessions[ps_name]["services"]:
                ttk.Label(bill_frame, text=f"• {service['name'].title()}: ${int(service['price']):,} ({service['time']})").pack(anchor=tk.W)
        else:
            ttk.Label(bill_frame, text="No additional services").pack(anchor=tk.W)
        
        ttk.Label(bill_frame, text=f"Services Total: ${int(service_cost):,}").pack(anchor=tk.W, pady=(5, 0))
        
        ttk.Separator(bill_frame, orient='horizontal').pack(fill=tk.X, pady=10)
        
        # Total section (will be updated dynamically)
        self.bill_total_frame = ttk.Frame(bill_frame)
        self.bill_total_frame.pack(anchor=tk.W)
        
        # Initial total calculation
        self.update_bill_total(bill_frame, duration_hours, service_cost)
        
        # Buttons
        button_frame = ttk.Frame(bill_frame)
        button_frame.pack(pady=20)
        
        ttk.Button(button_frame, text="Save to Database", 
                  command=lambda: self.save_bill_to_database(ps_name, duration, service_cost, bill_window)).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(button_frame, text="Close Without Saving", 
                  command=lambda: self.close_session(ps_name, bill_window)).pack(side=tk.LEFT, padx=5)

    def update_bill_total(self, bill_frame, duration_hours, service_cost):
        """Update the total cost based on offer selection"""
        # Clear previous total display
        for widget in self.bill_total_frame.winfo_children():
            widget.destroy()
        
        # Calculate PlayStation cost based on offer selection
        if self.apply_offer_var.get():
            ps_cost = self.current_offer["offer_cost"]
            savings = self.current_offer["savings"]
            
            ttk.Label(self.bill_total_frame, text=f"PlayStation Cost: ${int(ps_cost):,} (with offer)", 
                     font=("Arial", 12, "bold"), foreground="green").pack(anchor=tk.W)
            if savings > 0:
                ttk.Label(self.bill_total_frame, text=f"Savings Applied: ${int(savings):,}", 
                         font=("Arial", 10), foreground="red").pack(anchor=tk.W)
        else:
            ps_cost = self.current_offer["normal_cost"]
            ttk.Label(self.bill_total_frame, text=f"PlayStation Cost: ${int(ps_cost):,} (normal rate)", 
                     font=("Arial", 12, "bold")).pack(anchor=tk.W)
        
        # Calculate and display total
        total_cost = ps_cost + service_cost
        ttk.Label(self.bill_total_frame, text=f"TOTAL: ${int(total_cost):,}", 
                 font=("Arial", 14, "bold")).pack(anchor=tk.W, pady=(10, 0))

    def save_bill_to_database(self, ps_name, duration, service_cost, bill_window):
        """Save session data to Excel database with offer applied"""
        # Calculate final PlayStation cost based on offer selection
        if self.apply_offer_var.get():
            ps_cost = self.current_offer["offer_cost"]
        else:
            ps_cost = self.current_offer["normal_cost"]
        
        total_cost = ps_cost + service_cost
        
        # Confirmation dialog
        result = messagebox.askyesno("Confirm Save", 
                                    f"Save this session to database?\n\nTotal: ${int(total_cost):,}")
        if not result:
            return
        
        try:
            # Read existing data
            df = pd.read_excel(self.db_file)
            
            # Prepare services string
            services_str = ", ".join([f"{s['name']}(${int(s['price']):,})" for s in self.sessions[ps_name]["services"]])
            if not services_str:
                services_str = "None"
            
            # Format duration as HH:MM
            hours = int(duration // 3600)
            minutes = int((duration % 3600) // 60)
            duration_formatted = f"{hours:02d}:{minutes:02d}"
            
            # Create new row
            new_row = {
                'Date': datetime.now().strftime('%Y-%m-%d'),
                'Time': datetime.now().strftime('%H:%M:%S'),
                'PlayStation': ps_name,
                'Customer': "N/A",
                'Duration_Hours': duration_formatted,
                'PS_Cost': round(ps_cost, 2),
                'Services': services_str,
                'Service_Cost': round(service_cost, 2),
                'Total_Cost': round(total_cost, 2)
            }
            
            # Add to dataframe
            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
            
            # Save to Excel
            df.to_excel(self.db_file, index=False)
            
            messagebox.showinfo("Success", "Session saved to database!")
            self.close_session(ps_name, bill_window)
            
            # Auto-refresh database tab if it exists
            if hasattr(self, 'tree'):
                self.refresh_database()
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save to database: {str(e)}")
    
    def close_session(self, ps_name, bill_window):
        """Close session and reset PlayStation"""
        # Reset session
        self.sessions[ps_name] = {
            "active": False,
            "start_time": None,
            "customer_name": "",
            "services": []
        }
        
        # Reset UI
        self.ps_frames[ps_name]["status_label"].config(text="Available", foreground="green")
        self.ps_frames[ps_name]["timer_label"].config(text="00:00:00")
        self.ps_frames[ps_name]["apply_btn"].config(state="normal")
        self.ps_frames[ps_name]["services_listbox"].delete(0, tk.END)
        
        bill_window.destroy()
    
    def download_excel(self):
        """Download Excel database file"""
        # Confirmation dialog
        result = messagebox.askyesno("Confirm Export", "Export database to Excel file?")
        if not result:
            return
        
        try:
            if not os.path.exists(self.db_file):
                messagebox.showwarning("Warning", "No database file found!")
                return
            
            # Ask user where to save
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                title="Save Gaming Lounge Database"
            )
            
            if file_path:
                # Ensure .xlsx extension
                if not file_path.endswith('.xlsx'):
                    file_path += '.xlsx'
                
                # Copy database file to chosen location
                df = pd.read_excel(self.db_file)
                df.to_excel(file_path, index=False)
                messagebox.showinfo("Success", f"Database exported to:\n{file_path}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export database: {str(e)}")

    def update_timer(self):
        for ps_name, session in self.sessions.items():
            if session["active"]:
                duration = time.time() - session["start_time"]
                hours = int(duration // 3600)
                minutes = int((duration % 3600) // 60)
                seconds = int(duration % 60)
                
                timer_text = f"{hours:02d}:{minutes:02d}:{seconds:02d}"
                self.ps_frames[ps_name]["timer_label"].config(text=timer_text)
                
                # Calculate PlayStation cost with offers
                duration_hours = duration / 3600
                current_ps_cost = self.calculate_ps_cost(duration_hours)
                
                # Add services cost
                services_cost = sum(service["price"] for service in session["services"])
                total_current_cost = current_ps_cost + services_cost
                
                # Format cost with commas and no decimals
                cost_formatted = f"{int(total_current_cost):,}"
                self.ps_frames[ps_name]["cost_label"].config(text=f"Cost: ${cost_formatted}")
            else:
                # Reset cost when not active
                self.ps_frames[ps_name]["cost_label"].config(text="Cost: $0")
        
        self.root.after(1000, self.update_timer)

    def calculate_ps_cost(self, duration_hours):
        """Calculate PlayStation cost with offers applied"""
        if not self.config["offers"]["enabled"]:
            return duration_hours * self.config["playstation_rate"]
        
        base_rate = self.config["playstation_rate"]
        
        if duration_hours >= 3:
            # 3+ hours: all hours at 4666 rate
            return duration_hours * self.config["offers"]["3_hour_rate"]
        elif duration_hours >= 2:
            # 2+ hours: all hours at 5000 rate
            return duration_hours * self.config["offers"]["2_hour_rate"]
        else:
            # Less than 2 hours: normal rate
            return duration_hours * base_rate

    def setup_database_tab(self):
        """Setup the live database viewer tab"""
        # Main frame
        db_frame = ttk.Frame(self.database_tab, padding="10")
        db_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title and controls
        header_frame = ttk.Frame(db_frame)
        header_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(header_frame, text="Database Viewer", 
                 font=("Arial", 16, "bold")).pack(side=tk.LEFT)
        
        # Date selection frame
        date_frame = ttk.Frame(header_frame)
        date_frame.pack(side=tk.LEFT, padx=(20, 0))
        
        ttk.Label(date_frame, text="Select Date:").pack(side=tk.LEFT, padx=(0, 5))
        
        # Date selection dropdowns
        today = datetime.now()
        
        # Year dropdown
        self.year_var = tk.StringVar(value=str(today.year))
        year_combo = ttk.Combobox(date_frame, textvariable=self.year_var, width=6, state="readonly")
        year_combo['values'] = [str(year) for year in range(2020, 2030)]
        year_combo.pack(side=tk.LEFT, padx=2)
        
        # Month dropdown
        self.month_var = tk.StringVar(value=f"{today.month:02d}")
        month_combo = ttk.Combobox(date_frame, textvariable=self.month_var, width=4, state="readonly")
        month_combo['values'] = [f"{month:02d}" for month in range(1, 13)]
        month_combo.pack(side=tk.LEFT, padx=2)
        
        # Day dropdown
        self.day_var = tk.StringVar(value=f"{today.day:02d}")
        day_combo = ttk.Combobox(date_frame, textvariable=self.day_var, width=4, state="readonly")
        day_combo['values'] = [f"{day:02d}" for day in range(1, 32)]
        day_combo.pack(side=tk.LEFT, padx=2)
        
        # Load date button
        ttk.Button(date_frame, text="Load Date", 
                  command=self.load_date_data).pack(side=tk.LEFT, padx=5)
        
        # Show all data button
        ttk.Button(date_frame, text="Show All", 
                  command=self.show_all_data).pack(side=tk.LEFT, padx=5)
        
        # Control buttons
        button_frame = ttk.Frame(header_frame)
        button_frame.pack(side=tk.RIGHT)
        
        ttk.Button(button_frame, text="Refresh", 
                  command=self.refresh_database).pack(side=tk.LEFT, padx=2)
        ttk.Button(button_frame, text="Delete Selected", 
                  command=self.delete_selected_row).pack(side=tk.LEFT, padx=2)
        ttk.Button(button_frame, text="Export Excel", 
                  command=self.download_excel).pack(side=tk.LEFT, padx=2)
        
        # Daily summary frame
        summary_frame = ttk.LabelFrame(db_frame, text="Daily Summary", padding="10")
        summary_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Summary labels
        summary_info_frame = ttk.Frame(summary_frame)
        summary_info_frame.pack(fill=tk.X)
        
        self.summary_date_label = ttk.Label(summary_info_frame, text="Date: Today", 
                                           font=("Arial", 12, "bold"))
        self.summary_date_label.pack(side=tk.LEFT)
        
        self.summary_total_label = ttk.Label(summary_info_frame, text="Total Revenue: $0.00", 
                                            font=("Arial", 12, "bold"), foreground="green")
        self.summary_total_label.pack(side=tk.RIGHT)
        
        # Breakdown frame
        breakdown_frame = ttk.Frame(summary_frame)
        breakdown_frame.pack(fill=tk.X, pady=(5, 0))
        
        self.summary_ps_label = ttk.Label(breakdown_frame, text="PlayStation: $0.00")
        self.summary_ps_label.pack(side=tk.LEFT)
        
        self.summary_services_label = ttk.Label(breakdown_frame, text="Services: $0.00")
        self.summary_services_label.pack(side=tk.LEFT, padx=(20, 0))
        
        self.summary_sessions_label = ttk.Label(breakdown_frame, text="Sessions: 0")
        self.summary_sessions_label.pack(side=tk.RIGHT)
        
        # Treeview for data display
        columns = ['Date', 'Time', 'PlayStation', 'Customer', 'Duration_Hours', 
                   'PS_Cost', 'Services', 'Service_Cost', 'Total_Cost']
        
        self.tree = ttk.Treeview(db_frame, columns=columns, show='headings', height=20)
        
        # Configure columns
        for col in columns:
            if col == 'Duration_Hours':
                self.tree.heading(col, text='Duration (HH:MM)')
            else:
                self.tree.heading(col, text=col.replace('_', ' '))
            
            if col in ['PS_Cost', 'Service_Cost', 'Total_Cost']:
                self.tree.column(col, width=80, anchor='center')
            elif col in ['Date', 'Time']:
                self.tree.column(col, width=100, anchor='center')
            elif col == 'PlayStation':
                self.tree.column(col, width=80, anchor='center')
            elif col == 'Duration_Hours':
                self.tree.column(col, width=100, anchor='center')
            else:
                self.tree.column(col, width=120)
        
        # Scrollbars
        v_scrollbar = ttk.Scrollbar(db_frame, orient=tk.VERTICAL, command=self.tree.yview)
        h_scrollbar = ttk.Scrollbar(db_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # Pack treeview and scrollbars
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Load initial data (today's data)
        self.refresh_database()

    def load_date_data(self):
        """Load data for selected date"""
        try:
            selected_date = f"{self.year_var.get()}-{self.month_var.get()}-{self.day_var.get()}"
            # Validate date
            datetime.strptime(selected_date, '%Y-%m-%d')
            self.refresh_database(filter_date=selected_date)
        except ValueError:
            messagebox.showerror("Error", "Invalid date selected!")

    def show_all_data(self):
        """Show all data regardless of date"""
        self.refresh_database(show_all=True)

    def refresh_database(self, filter_date=None, show_all=False):
        """Refresh the database view with optional date filtering"""
        try:
            # Clear existing data
            for item in self.tree.get_children():
                self.tree.delete(item)
            
            # Load data from Excel
            if os.path.exists(self.db_file):
                df = pd.read_excel(self.db_file)
                
                # Filter by date if specified
                if filter_date and not show_all:
                    df_filtered = df[df['Date'] == filter_date]
                    display_date = filter_date
                elif not show_all:
                    # Default to today's date
                    today = datetime.now().strftime('%Y-%m-%d')
                    df_filtered = df[df['Date'] == today]
                    display_date = today
                else:
                    # Show all data
                    df_filtered = df
                    display_date = "All Dates"
                
                # Insert filtered data into treeview
                for index, row in df_filtered.iterrows():
                    values = [str(row[col]) if pd.notna(row[col]) else "" for col in df.columns]
                    self.tree.insert("", tk.END, values=values)
                
                # Update summary
                self.update_daily_summary(df_filtered, display_date)
                
                if not show_all:
                    messagebox.showinfo("Success", f"Database refreshed for {display_date}!")
                else:
                    messagebox.showinfo("Success", "Database refreshed - showing all data!")
            else:
                # Reset summary if no data
                self.update_daily_summary(pd.DataFrame(), "No Data")
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to refresh database: {str(e)}")

    def update_daily_summary(self, df, display_date):
        """Update the daily summary section"""
        if df.empty:
            self.summary_date_label.config(text=f"Date: {display_date}")
            self.summary_total_label.config(text="Total Revenue: $0")
            self.summary_ps_label.config(text="PlayStation: $0")
            self.summary_services_label.config(text="Services: $0")
            self.summary_sessions_label.config(text="Sessions: 0")
            return
        
        # Calculate totals
        total_revenue = df['Total_Cost'].sum() if 'Total_Cost' in df.columns else 0
        ps_revenue = df['PS_Cost'].sum() if 'PS_Cost' in df.columns else 0
        services_revenue = df['Service_Cost'].sum() if 'Service_Cost' in df.columns else 0
        total_sessions = len(df)
        
        # Format numbers without decimals and with commas
        total_formatted = f"{int(total_revenue):,}"
        ps_formatted = f"{int(ps_revenue):,}"
        services_formatted = f"{int(services_revenue):,}"
        
        # Update labels
        self.summary_date_label.config(text=f"Date: {display_date}")
        self.summary_total_label.config(text=f"Total Revenue: ${total_formatted}")
        self.summary_ps_label.config(text=f"PlayStation: ${ps_formatted}")
        self.summary_services_label.config(text=f"Services: ${services_formatted}")
        self.summary_sessions_label.config(text=f"Sessions: {total_sessions}")

    def delete_selected_row(self):
        """Delete selected row from database"""
        selected_item = self.tree.selection()
        
        if not selected_item:
            messagebox.showwarning("Warning", "Please select a row to delete")
            return
        
        # Confirmation dialog
        result = messagebox.askyesno("Confirm Delete", 
                                    "Are you sure you want to delete the selected record?\n\nThis action cannot be undone!")
        
        if not result:
            return
        
        try:
            # Get the index of selected item
            item_index = self.tree.index(selected_item[0])
            
            # Load dataframe
            df = pd.read_excel(self.db_file)
            
            # Remove the row
            df = df.drop(df.index[item_index]).reset_index(drop=True)
            
            # Save back to Excel
            df.to_excel(self.db_file, index=False)
            
            # Refresh the view
            self.refresh_database()
            
            messagebox.showinfo("Success", "Record deleted successfully!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to delete record: {str(e)}")

    def add_service_only(self, service):
        """Add service to current services-only order"""
        # Show confirmation dialog
        result = messagebox.askyesno("Confirm Service", 
                                   f"Add {service.title()} (${self.config['services'][service]}) to current order?")
        
        if result:
            service_entry = {
                "name": service,
                "price": self.config["services"][service],
                "time": datetime.now().strftime("%H:%M:%S")
            }
            
            self.services_only_list.append(service_entry)
            self.update_current_order_display()
            messagebox.showinfo("Success", f"{service.title()} added to current order")

    def update_current_order_display(self):
        """Update the current order listbox display"""
        self.current_order_listbox.delete(0, tk.END)
        
        for service in self.services_only_list:
            display_text = f"{service['name'].title()} - ${service['price']} ({service['time']})"
            self.current_order_listbox.insert(tk.END, display_text)

    def add_to_pending_orders(self):
        """Add current order to pending orders"""
        if not self.services_only_list:
            messagebox.showwarning("Warning", "No services in current order")
            return
        
        customer_name = self.services_customer_var.get().strip()
        if not customer_name:
            messagebox.showwarning("Warning", "Please enter customer name")
            return
        
        # Calculate total
        total_cost = sum(service["price"] for service in self.services_only_list)
        
        # Confirmation dialog
        result = messagebox.askyesno("Confirm Add to Pending", 
                                   f"Add order for {customer_name} to pending orders?\n\nTotal: ${total_cost:.2f}")
        
        if result:
            # Create pending order
            pending_order = {
                "customer": customer_name,
                "services": self.services_only_list.copy(),
                "total": total_cost,
                "time_added": datetime.now().strftime("%H:%M:%S")
            }
            
            self.pending_orders.append(pending_order)
            self.update_pending_orders_display()
            
            # Clear current order
            self.services_only_list.clear()
            self.services_customer_var.set("")
            self.update_current_order_display()
            
            messagebox.showinfo("Success", f"Order for {customer_name} added to pending orders")

    def update_pending_orders_display(self):
        """Update the pending orders treeview"""
        # Clear existing items
        for item in self.pending_tree.get_children():
            self.pending_tree.delete(item)
        
        # Add pending orders
        for order in self.pending_orders:
            services_text = ", ".join([f"{s['name'].title()}(${s['price']})" for s in order["services"]])
            values = [
                order["customer"],
                services_text,
                f"${order['total']:.2f}",
                order["time_added"]
            ]
            self.pending_tree.insert("", tk.END, values=values)

    def generate_bill_for_pending(self):
        """Generate bill for selected pending order"""
        selected_item = self.pending_tree.selection()
        
        if not selected_item:
            messagebox.showwarning("Warning", "Please select a pending order")
            return
        
        # Get selected order index
        item_index = self.pending_tree.index(selected_item[0])
        order = self.pending_orders[item_index]
        
        # Confirmation dialog
        result = messagebox.askyesno("Confirm Bill Generation", 
                                   f"Generate bill for {order['customer']}?\n\nTotal: ${order['total']:.2f}")
        
        if result:
            self.show_pending_services_bill(order, item_index)

    def show_pending_services_bill(self, order, order_index):
        """Show bill for pending services order"""
        bill_window = tk.Toplevel(self.root)
        bill_window.title("Services Bill")
        bill_window.geometry("400x500")
        
        # Bill content
        bill_frame = ttk.Frame(bill_window, padding="20")
        bill_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(bill_frame, text="SERVICES BILL", 
                 font=("Arial", 16, "bold")).pack(pady=(0, 20))
        
        ttk.Label(bill_frame, text=f"Customer: {order['customer']}").pack(anchor=tk.W)
        ttk.Label(bill_frame, text=f"Date: {datetime.now().strftime('%Y-%m-%d %H:%M')}").pack(anchor=tk.W)
        ttk.Label(bill_frame, text=f"Order Time: {order['time_added']}").pack(anchor=tk.W)
        
        ttk.Separator(bill_frame, orient='horizontal').pack(fill=tk.X, pady=10)
        
        # Services
        ttk.Label(bill_frame, text="Services Ordered:", font=("Arial", 12, "bold")).pack(anchor=tk.W)
        
        for service in order["services"]:
            ttk.Label(bill_frame, text=f"• {service['name'].title()}: ${service['price']} ({service['time']})").pack(anchor=tk.W)
        
        ttk.Separator(bill_frame, orient='horizontal').pack(fill=tk.X, pady=10)
        
        # Total
        ttk.Label(bill_frame, text=f"TOTAL: ${order['total']:.2f}", 
                 font=("Arial", 14, "bold")).pack(anchor=tk.W)
        
        # Buttons
        button_frame = ttk.Frame(bill_frame)
        button_frame.pack(pady=20)
        
        ttk.Button(button_frame, text="Save & Complete Order", 
                  command=lambda: self.save_pending_to_database(order, order_index, bill_window)).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(button_frame, text="Close Without Saving", 
                  command=bill_window.destroy).pack(side=tk.LEFT, padx=5)

    def save_pending_to_database(self, order, order_index, bill_window):
        """Save pending order to database and remove from pending"""
        result = messagebox.askyesno("Confirm Save", 
                                   f"Save and complete order for {order['customer']}?\n\nTotal: ${order['total']:.2f}")
        if not result:
            return
        
        try:
            # Read existing data
            df = pd.read_excel(self.db_file)
            
            # Prepare services string
            services_str = ", ".join([f"{s['name']}(${s['price']})" for s in order["services"]])
            
            # Create new row
            new_row = {
                'Date': datetime.now().strftime('%Y-%m-%d'),
                'Time': datetime.now().strftime('%H:%M:%S'),
                'PlayStation': "Services Only",
                'Customer': order["customer"],
                'Duration_Hours': "00:00",
                'PS_Cost': 0.00,
                'Services': services_str,
                'Service_Cost': round(order["total"], 2),
                'Total_Cost': round(order["total"], 2)
            }
            
            # Add to dataframe
            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
            
            # Save to Excel
            df.to_excel(self.db_file, index=False)
            
            # Remove from pending orders
            del self.pending_orders[item_index]
            self.update_pending_orders_display()
            
            messagebox.showinfo("Success", f"Order for {order['customer']} completed and saved!")
            bill_window.destroy()
            
            # Auto-refresh database tab if it exists
            if hasattr(self, 'tree'):
                self.refresh_database()
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save order: {str(e)}")

    def remove_pending_order(self):
        """Remove selected pending order"""
        selected_item = self.pending_tree.selection()
        
        if not selected_item:
            messagebox.showwarning("Warning", "Please select a pending order to remove")
            return
        
        # Get selected order
        item_index = self.pending_tree.index(selected_item[0])
        order = self.pending_orders[item_index]
        
        result = messagebox.askyesno("Confirm Removal", 
                                   f"Remove pending order for {order['customer']}?\n\nThis action cannot be undone!")
        
        if result:
            del self.pending_orders[item_index]
            self.update_pending_orders_display()
            messagebox.showinfo("Success", "Pending order removed")

    def add_more_services_to_pending(self):
        """Add more services to selected pending order"""
        selected_item = self.pending_tree.selection()
        
        if not selected_item:
            messagebox.showwarning("Warning", "Please select a pending order")
            return
        
        # Get selected order
        item_index = self.pending_tree.index(selected_item[0])
        order = self.pending_orders[item_index]
        
        # Load order into current order for editing
        self.services_only_list = order["services"].copy()
        self.services_customer_var.set(order["customer"])
        self.update_current_order_display()
        
        # Remove from pending (will be re-added when "Add to Pending" is clicked)
        del self.pending_orders[item_index]
        self.update_pending_orders_display()
        
        messagebox.showinfo("Info", f"Order for {order['customer']} loaded for editing.\nAdd more services and click 'Add to Pending Orders' when done.")

    def remove_service_only(self):
        """Remove selected service from current order"""
        selection = self.current_order_listbox.curselection()
        
        if not selection:
            messagebox.showwarning("Warning", "Please select a service to remove")
            return
        
        index = selection[0]
        service_name = self.services_only_list[index]["name"]
        
        result = messagebox.askyesno("Confirm Removal", 
                                   f"Remove {service_name.title()} from current order?")
        
        if result:
            del self.services_only_list[index]
            self.update_current_order_display()

    def clear_current_order(self):
        """Clear current order"""
        if not self.services_only_list:
            messagebox.showinfo("Info", "No services in current order")
            return
        
        result = messagebox.askyesno("Confirm Clear", "Clear current order?")
        
        if result:
            self.services_only_list.clear()
            self.services_customer_var.set("")
            self.update_current_order_display()
            messagebox.showinfo("Success", "Current order cleared")

    def run(self):
        """Start the application"""
        self.root.mainloop()

if __name__ == "__main__":
    app = GamingLoungeManager()
    app.run()
