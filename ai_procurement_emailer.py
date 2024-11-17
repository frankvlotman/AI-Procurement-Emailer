import tkinter as tk
from tkinter import scrolledtext, messagebox, Toplevel, Label, Text, Button, ttk
from langchain_openai import ChatOpenAI
from langchain.prompts.prompt import PromptTemplate
from langchain.chains import LLMChain
import os
from PIL import Image
import win32com.client as win32
from docx import Document

# Define the path for the blank icon
icon_path = 'C:\\Users\\Frank\\Desktop\\blank.ico'

# Create a blank (transparent) ICO file if it doesn't exist
def create_blank_ico(path):
    size = (16, 16)  # Size of the icon
    image = Image.new("RGBA", size, (255, 255, 255, 0))  # Transparent image
    image.save(path, format="ICO")

# Create the blank ICO file
create_blank_ico(icon_path)

# Set the OpenAI API key
os.environ["OPENAI_API_KEY"] = "API Key Here"

# Initialize the OpenAI model
llm = ChatOpenAI(model_name="gpt-4o-mini")

# Predefined email scenarios with added "Chase Up Response"
email_scenarios = {
    "RFQ: Current Supplier": "Write a professional request for quotation email to {supplier_name} for the following products: {product_description}. Include delivery expectations and ask for price and lead time.",
    "RFQ: New Supplier": "I hope this message finds you well. My name is x, and I am the Procurement Officer at x Company. We are currently exploring options for {product_description}, and we came across your company as a potential supplier. Please provide a quotation including pricing, delivery times, and other relevant details.",
    "Order Confirmation": "Write a polite order confirmation email to {supplier_name} for the following order: {order_details}. Confirm the delivery date and terms.",
    "Credit Note Request": "Request a credit note from {supplier_name} for the following invoice: {invoice_details}. Explain the reason for the request, whether due to incorrect charges or a return of goods.",
    "Supplier Feedback": "Provide constructive feedback to {supplier_name} on the quality of the latest delivery of {product_description}. Request improvements.",
    "Earlier Delivery": "Request an earlier delivery date from {supplier_name} for the urgent requirement of {product_description}. Explain the urgency and request for a sooner delivery date.",
    "Chase Up Response": "Write a polite follow-up email to {supplier_name} to chase up their response regarding the earlier communication about {product_description}. Kindly ask for a status update."
}

# Define the Casual Human prompt
CASUAL_HUMAN_PROMPT = (
    "Use a human voice with a friendly tone that isn’t colloquial. "
    "Use short sentences and simple words. Remove academic language, transition phrases, and corporate jargon. "
    "Additionally, check for common AI writing patterns such as repeated sentence structures, overused transitions, "
    "unnecessary clarifications, and academic phrasing. Rework these elements to sound more natural and human-like. "
    "Keep the key points but strip away any unnecessary words."
)

# Create a prompt template with placeholders
prompt_template = PromptTemplate(
    input_variables=["scenario", "tone", "supplier_name", "product_description", "order_details", "invoice_details", "additional_prompt"],
    template="Scenario: {scenario}\nTone: {tone}\nSupplier: {supplier_name}\nDetails: {product_description}\nOrder: {order_details}\nInvoice: {invoice_details}\nAdditional Prompt: {additional_prompt}\nResponse:"
)

# Create the LLM chain
chain1 = LLMChain(prompt=prompt_template, llm=llm)

# Function to get response from OpenAI and display it in a new window
def get_response_and_display():
    user_prompt = prompt_entry.get("1.0", tk.END).strip()  # Get any additional prompt
    selected_scenario = email_type.get()  # Get selected email scenario
    selected_tone = email_tone.get()  # Get selected email tone
    supplier_name = supplier_entry.get().strip()  # Supplier name input
    product_description = product_entry.get().strip()  # Product description input
    order_details = order_entry.get().strip()  # Order details input
    invoice_details = invoice_entry.get().strip()  # Invoice details input
    
    # If additional prompt is blank, set it to an empty string to avoid issues
    if not user_prompt:
        user_prompt = ""
    
    # Append Casual Human prompt if selected
    if selected_tone == "Casual Human":
        if user_prompt:
            user_prompt += "\n" + CASUAL_HUMAN_PROMPT
        else:
            user_prompt = CASUAL_HUMAN_PROMPT

    if selected_scenario:  # Only check for the scenario, not additional prompt
        status_label.config(text="Fetching response...")
        root.update_idletasks()
        try:
            # Pass the additional prompt and other fields into the LLM chain
            response = chain1({
                "scenario": email_scenarios[selected_scenario],
                "tone": selected_tone,
                "supplier_name": supplier_name,
                "product_description": product_description,
                "order_details": order_details,
                "invoice_details": invoice_details,
                "additional_prompt": user_prompt  # Include additional prompt
            })
            content = response["text"]
            status_label.config(text="Response received, displaying...")
            root.update_idletasks()
            
            # Create a new window for displaying the response
            response_window = Toplevel(root)
            response_window.title("Response")
            response_window.geometry("600x600")
            
            # Set the window icon
            response_window.iconbitmap(icon_path)

            # Display the response in the new window
            response_label = Label(response_window, text="Response:", font=("Arial", 14))
            response_label.pack(pady=10)
            
            response_text = Text(response_window, wrap=tk.WORD, font=("Arial", 12))
            response_text.insert(tk.END, content)
            response_text.config(state=tk.DISABLED)
            response_text.pack(pady=10, padx=10, expand=True, fill=tk.BOTH)
            
            # Function to download response to docx
            def download_to_docx():
                filepath = "C:\\Users\\Frank\\Desktop\\response.docx"
                doc = Document()
                doc.add_paragraph(content)
                doc.save(filepath)
                messagebox.showinfo("Saved", f"Response saved to {filepath}")
            
            # Function to send the response via email
            def send_email():
                try:
                    outlook = win32.Dispatch('outlook.application')
                    mail = outlook.CreateItem(0)
                    mail.To = 'email address here'
                    mail.Subject = 'Procurement Email'
                    mail.Body = content
                    mail.Send()
                    messagebox.showinfo("Email Sent", "Email sent successfully.")
                except Exception as e:
                    messagebox.showerror("Error", f"An error occurred while sending email: {e}")
            
            # Buttons for downloading and sending email
            button_frame = tk.Frame(response_window)
            button_frame.pack(pady=10)

            download_button = Button(button_frame, text="Download", command=download_to_docx, bg='#d0e8f1')
            download_button.pack(side=tk.LEFT, padx=10)

            email_button = Button(button_frame, text="Email", command=send_email, bg='#d0e8f1')
            email_button.pack(side=tk.RIGHT, padx=10)
            
            status_label.config(text="Response displayed")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")
            status_label.config(text="Error")
    else:
        messagebox.showwarning("Input Error", "Please select an email type.")
        status_label.config(text="")

# Setting up the GUI
root = tk.Tk()
root.title("Procurement Email Assistant")

# Set the window icon
root.iconbitmap(icon_path)

# Email Scenario Selection
scenario_label = tk.Label(root, text="Select Email Type:")
scenario_label.pack(pady=5)
email_type = tk.StringVar(root)
email_dropdown = ttk.Combobox(root, textvariable=email_type, state="readonly")
email_dropdown['values'] = list(email_scenarios.keys())
email_dropdown.pack(pady=5)

# Tone Selection
tone_label = tk.Label(root, text="Select Email Tone:")
tone_label.pack(pady=5)
email_tone = tk.StringVar(root)
tone_dropdown = ttk.Combobox(root, textvariable=email_tone, state="readonly")
tone_dropdown['values'] = ["Formal", "Semi-Formal", "Casual", "Casual Human"]
tone_dropdown.pack(pady=5)

# Supplier Name Entry
supplier_label = tk.Label(root, text="Supplier Name:")
supplier_label.pack(pady=5)
supplier_entry = tk.Entry(root, width=50)
supplier_entry.pack(pady=5)

# Product Description Entry
product_label = tk.Label(root, text="Product Description:")
product_label.pack(pady=5)
product_entry = tk.Entry(root, width=50)
product_entry.pack(pady=5)

# Order Details Entry
order_label = tk.Label(root, text="Order Details:")
order_label.pack(pady=5)
order_entry = tk.Entry(root, width=50)
order_entry.pack(pady=5)

# Invoice Details Entry
invoice_label = tk.Label(root, text="Invoice Details:")
invoice_label.pack(pady=5)
invoice_entry = tk.Entry(root, width=50)
invoice_entry.pack(pady=5)

# Prompt label and text entry
prompt_label = tk.Label(root, text="Enter any additional prompt (optional):")
prompt_label.pack(pady=5)
prompt_entry = scrolledtext.ScrolledText(root, height=5, width=50)
prompt_entry.pack(pady=5)

# Status label
status_label = tk.Label(root, text="", fg="blue")
status_label.pack(pady=5)

# Frame to hold the buttons
button_frame = tk.Frame(root)
button_frame.pack(pady=20)

# Submit button
submit_button = tk.Button(button_frame, text="Submit", command=get_response_and_display, bg='#d0e8f1')
submit_button.pack(side=tk.LEFT, padx=10)

# Start the GUI event loop
root.mainloop()
