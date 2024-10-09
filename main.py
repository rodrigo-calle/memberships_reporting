import os
import tkinter as tk
from tkinter import ttk
import firebase_admin
from firebase_admin import credentials, firestore
import pandas as pd
from datetime import datetime

def export_data():
    export_button.config(text="Exportando...", state="disabled", style="TButton")
    root.after(100, perform_export)

def perform_export():
    try:
        cred = credentials.Certificate('private/firebase_keys.json')
        if not firebase_admin._apps:
            firebase_admin.initialize_app(cred)

        db = firestore.client()

        no_subscriptions = []
        small_subscriptions = []
        large_subscriptions = []
        no_subscription_statuses = []
        small_subscription_statuses = []
        large_subscription_statuses = []

        doc_ref = db.collection('site_configs')
        docs = doc_ref.stream()

        for doc in docs:
            doc_id = doc.id
            doc_data = doc.to_dict()

            if not doc_data:
                print(doc_id)
                continue

            if 'settings' not in doc_data:
                print('Site settings not found: ', doc_id)
                continue

            id_required = doc_data['settings'].get('id_required', False)
            index_index_in_search_engines = doc_data['settings'].get('index_index_in_search_engines', False)

            notifications = doc_data['settings'].get('notifications', {})
            cancelled_membership = notifications.get('cancelled_memberships', False)
            new_subscription = notifications.get('new_memberships', False)
            transaction_declined = notifications.get('transaction_declined', False)
            three_days_before_renewal = doc_data['settings'].get('user_notifications', {}).get('three_days_before_renewal', False)

            is_active = (id_required or index_index_in_search_engines or
                         cancelled_membership or new_subscription or
                         transaction_declined or three_days_before_renewal)

            site_status = "Active" if is_active else "Not Active"

            memberships = db.collection('memberships').where('site', '==', doc_id).stream()
            membership_count = sum(1 for _ in memberships)

            if membership_count == 0:
                no_subscriptions.append(doc_id)
                no_subscription_statuses.append(site_status)
            elif 1 <= membership_count <= 10:
                small_subscriptions.append(doc_id)
                small_subscription_statuses.append(site_status)
            else:
                large_subscriptions.append(doc_id)
                large_subscription_statuses.append(site_status)

        data_no_subscriptions = {
            'No Subscriptions': no_subscriptions,
            'Site Status': no_subscription_statuses,
        }

        data_small_subscriptions = {
            '1-10 Subscriptions': small_subscriptions,
            'Site Status': small_subscription_statuses,
        }

        data_large_subscriptions = {
            'More than 10 Subscriptions': large_subscriptions,
            'Site Status': large_subscription_statuses,
        }

        # Guardar el archivo
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f'site_subscriptions_report_{timestamp}.xlsx'

        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            if no_subscriptions:
                pd.DataFrame(data_no_subscriptions).to_excel(writer, sheet_name='No Subscriptions', index=False)
            if small_subscriptions:
                pd.DataFrame(data_small_subscriptions).to_excel(writer, sheet_name='1-10 Subscriptions', index=False)
            if large_subscriptions:
                pd.DataFrame(data_large_subscriptions).to_excel(writer, sheet_name='More than 10 Subscriptions', index=False)

        success_message.set(f"Data exported successfully to {filename}!")

    except Exception as e:
        success_message.set(f"An error occurred: {str(e)}")

    finally:
        export_button.config(text="Export to Excel", state="normal", style="TButton")

def create_gui():
    global root, export_button
    root = tk.Tk()
    root.title("Firestore to Excel Exporter")

    ttk.Label(root, text="Press the button to export data to Excel").grid(row=0, column=0, padx=10, pady=10)

    style = ttk.Style()
    style.configure('Export.TButton', background='red', foreground='white')

    export_button = ttk.Button(root, text="Export to Excel", command=export_data)
    export_button.grid(row=1, column=0, padx=10, pady=10)

    global success_message
    success_message = tk.StringVar()
    ttk.Label(root, textvariable=success_message).grid(row=2, column=0, padx=10, pady=10)

    root.mainloop()

create_gui()
