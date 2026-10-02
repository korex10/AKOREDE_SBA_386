"""SBA 386 - Data Analysis and Visualization with Python

This Python script completes the skill-based assignment using the
woocommerce-product-export.csv dataset and the Section Two tasks from the
course PDF.
"""

import os
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from docx import Document
from docx.shared import Inches


# ------------------------------
# Configuration
# ------------------------------
BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "SBA_386_outputs"
OUTPUT_DIR.mkdir(exist_ok=True)

CSV_PATH = Path(r"c:\Users\USER\Downloads\woocommerce-product-export.csv")
DOC_PATH = BASE_DIR / "YOURNAME_386_DOC_SBA.docx"


# ------------------------------
# Helper functions
# ------------------------------
def save_chart(fig, filename):
    fig.tight_layout()
    fig.savefig(OUTPUT_DIR / filename, dpi=200)
    plt.close(fig)


def add_chart_to_doc(doc, title, image_filename):
    doc.add_heading(title, level=1)
    doc.add_picture(str(OUTPUT_DIR / image_filename), width=Inches(6.5))


# ------------------------------
# Section One - CSV analysis and plotting
# ------------------------------
print("\n=== SECTION ONE ===")

df = pd.read_csv(CSV_PATH)
print("1. Reading data from CSV")
print(df)
print("\n2. Displaying info() summary")
df.info()
print("\n3. Displaying descriptive statistics")
print(df.describe(include="all"))
print("\n4. First five rows")
print(df.head())
print("\n5. Last five rows")
print(df.tail())
print("\n6. month_number and total_profit columns")
print(df[["month_number", "total_profit"]])

# 7. Bar plot for total profit by month
fig, ax = plt.subplots(figsize=(10, 6))
ax.bar(df["month_number"], df["total_profit"], color="steelblue")
ax.set_xlabel("Month number")
ax.set_ylabel("Total profit")
ax.set_title("Company profit per month")
ax.grid(axis="y", linestyle="--", alpha=0.3)
save_chart(fig, "bar_total_profit.png")

# 8. Line plot for total profit by month
fig, ax = plt.subplots(figsize=(10, 6))
ax.plot(
    df["month_number"],
    df["total_profit"],
    linestyle=":",
    color="red",
    marker="o",
    markerfacecolor="black",
    markeredgecolor="black",
    linewidth=3,
    label="Total profit",
)
ax.set_xlabel("Month number")
ax.set_ylabel("Total profit")
ax.set_title("Company profit per month")
ax.legend(loc="lower right")
save_chart(fig, "line_total_profit.png")

# 9. Multi-line plot for each product sale quantity by month
product_columns = [
    "facecream",
    "facewash",
    "toothpaste",
    "bathingsoap",
    "shampoo",
    "moisturizer",
]
fig, ax = plt.subplots(figsize=(12, 7))
for col in product_columns:
    ax.plot(df["month_number"], df[col], marker="o", linewidth=1.8, label=col)
ax.set_xlabel("Month number")
ax.set_ylabel("Units sold")
ax.set_title("Product sales per month")
ax.legend(loc="upper right", bbox_to_anchor=(1.02, 1.0))
save_chart(fig, "multiline_product_sales.png")

# 10. Scatter plot for bathingsoap sales
fig, ax = plt.subplots(figsize=(10, 6))
ax.scatter(df["month_number"], df["bathingsoap"], color="darkorange")
ax.set_xlabel("Month number")
ax.set_ylabel("Bathingsoap")
ax.set_title("Bathingsoap sales per month")
ax.grid(True, linestyle="--")
save_chart(fig, "scatter_bathingsoap.png")


# ------------------------------
# Section Two - Additional plotting tasks
# ------------------------------
print("\n=== SECTION TWO ===")

# Task 1: date vs temperature
print("\nDate-wise Temperature data")
dates = ["25/12", "26/12", "27/12"]
temperatures = [8.5, 10.5, 6.8]
for d, t in zip(dates, temperatures):
    print(f"{d}: {t} C")

fig, ax = plt.subplots(figsize=(10, 6))
ax.plot(dates, temperatures, color="blue", marker="o", linewidth=2)
ax.set_xlabel("Date")
ax.set_ylabel("Temperature")
ax.set_title("Date-wise Temperature")
save_chart(fig, "date_temperature_line.png")

# Task 2: average weight vs average height
print("\nAverage weight with respect to average height")
height = [121.9, 124.5, 129.5, 134.6, 139.7, 147.3, 152.4, 157.5, 162.6]
weight = [19.7, 21.3, 23.5, 25.9, 28.5, 32.1, 35.7, 39.6, 43.2]

avg_height = np.mean(height)
avg_weight = np.mean(weight)
print(f"Average height = {avg_height}")
print(f"Average weight = {avg_weight}")

fig, ax = plt.subplots(figsize=(8, 6))
ax.plot(
    [avg_weight],
    [avg_height],
    linestyle="-.",
    color="green",
    marker="o",
    markersize=10,
    markerfacecolor="green",
    markeredgecolor="green",
    label="Average values",
)
ax.set_xlabel("Weight in kg.")
ax.set_ylabel("Height in cm.")
ax.set_title("Average weight with respect to average height")
ax.legend(loc="lower right")
save_chart(fig, "avg_weight_height.png")


# ------------------------------
# Create Word document with screenshots
# ------------------------------
print("\nCreating Word document with all chart screenshots...")
doc = Document()
doc.add_heading("SBA 386 - Data Analysis and Visualization with Python", 0)
doc.add_paragraph(
    "This document contains the required charts and analysis for Section One and Section Two."
)

chart_list = [
    ("1. Bar Plot - Company Profit Per Month", "bar_total_profit.png"),
    ("2. Line Plot - Company Profit Per Month", "line_total_profit.png"),
    ("3. Multi-line Product Sales Plot", "multiline_product_sales.png"),
    ("4. Scatter Plot - Bathingsoap Sales", "scatter_bathingsoap.png"),
    ("5. Date-wise Temperature", "date_temperature_line.png"),
    ("6. Average Weight with Respect to Average Height", "avg_weight_height.png"),
]

for title, filename in chart_list:
    add_chart_to_doc(doc, title, filename)

# Include a quick summary section
summary = doc.add_paragraph()
summary.add_run("Summary: ")
summary.add_run(
    "The dataset contains monthly product sales and profit totals. The required charts "
    "have been produced in line with the assignment instructions."
)

# Save the file to the spellings required by the assignment
# The instruction asks for .doc; python-docx creates a .docx file which is the current Word standard.
doc.save(DOC_PATH)
print(f"Word file saved to: {DOC_PATH}")
print(f"Charts saved to: {OUTPUT_DIR}")

print("\nCompleted successfully.")
