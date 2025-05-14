import streamlit as st
import pandas as pd
import plotly.express as px
from utils import load_and_clean_excel, save_to_excel, save_budget, load_budget

st.set_page_config(page_title="מעקב הוצאות", layout="wide")

# טען CSS
with open("pastel.css") as f:
    st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)

st.title("אפליקציית מעקב הוצאות חודשיות")

# שלב 1 – העלאת קבצים
st.header("1. העלאת קבצים")
col1, col2 = st.columns(2)
with col1:
    prev_file = st.file_uploader("בחרי קובץ הוצאות קודם", type=["xlsx"], key="prev")
with col2:
    new_file = st.file_uploader("בחרי קובץ חודש חדש", type=["xlsx"], key="new")

if new_file:
    df_new = load_and_clean_excel(new_file)
    df_prev = load_and_clean_excel(prev_file) if prev_file else pd.DataFrame(columns=df_new.columns)
    df_all = pd.concat([df_prev, df_new]).drop_duplicates().reset_index(drop=True)

    st.subheader("טבלת הוצאות")
    st.dataframe(df_all)

    # הורדת אקסל מאוחד
    st.subheader("2. הורדת קובץ מאוחד")
    excel_data = save_to_excel(df_all)
    st.download_button("הורידי קובץ אקסל מאוחד", data=excel_data, file_name="merged_expenses.xlsx")

    # גרף הוצאות לפי חודש
    st.subheader("3. גרף הוצאות לפי חודש")
    monthly = df_all.groupby("Month")["Amount"].sum().reset_index()
    monthly["3mo_avg"] = monthly["Amount"].rolling(window=3).mean()
    fig = px.bar(monthly, x="Month", y="Amount", title="סך הוצאות חודשי", text_auto=True)
    fig.add_scatter(x=monthly["Month"], y=monthly["3mo_avg"], mode="lines+markers", name="ממוצע 3 חודשים")
    st.plotly_chart(fig, use_container_width=True)

    # גרף עוגה לפי קטגוריה
    st.subheader("4. פילוח לפי קטגוריה")
    pie_data = df_all.groupby("Category")["Amount"].sum().reset_index()
    fig_pie = px.pie(pie_data, names="Category", values="Amount", title="הוצאות לפי קטגוריה")
    st.plotly_chart(fig_pie, use_container_width=True)

    # KPI – ניהול תקציבים
    st.subheader("5. ניהול תקציבים חודשיים")
    st.markdown("הכניסי יעד לכל קטגוריה. אפשר לשמור את התקציבים לקובץ ולהעלות אותם שוב בהפעלה הבאה.")
    
    current_month = df_new["Month"].iloc[0]
    previous_month = df_prev["Month"].max() if not df_prev.empty else None

    # טעינת תקציב קודם
    budget_data = load_budget()
    updated_budget = {}

    for cat in sorted(df_all["Category"].unique()):
        col1, col2, col3 = st.columns([3, 2, 3])
        curr_total = df_new[df_new["Category"] == cat]["Amount"].sum()
        prev_total = df_prev[df_prev["Category"] == cat]["Amount"].sum() if previous_month else 0
        budget = budget_data.get(cat, 0.0)

        with col1:
            new_budget = st.number_input(f"תקציב לקטגוריה: {cat}", value=budget, step=10.0)
        with col2:
            st.markdown(f"*החודש:* {curr_total:.0f} ₪")
        with col3:
            status = ""
            if new_budget > 0:
                if curr_total <= new_budget:
                    status = f"בתקציב ✅"
                else:
                    over = curr_total - new_budget
                    # בדיקה אם בחודש הקודם הייתה חריגה/חיסכון
                    if prev_total < new_budget:
                        room = new_budget - prev_total
                        status = f"חריגה של {over:.0f} ₪ (אך היה עודף של {room:.0f} ₪ בחודש שעבר)"
                    else:
                        status = f"חריגה של {over:.0f} ₪ ❌"
            st.markdown(f"<div style='color:#996699; font-weight:bold'>{status}</div>", unsafe_allow_html=True)

        updated_budget[cat] = new_budget

    # כפתור לשמירת התקציבים
    if st.button("שמרי תקציבים לקובץ"):
        save_budget(updated_budget)
        st.success("התקציב נשמר בהצלחה ל-budget.json")
