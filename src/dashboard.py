from __future__ import annotations

import pandas as pd
import plotly.express as px
import streamlit as st


def render_dashboard(df: pd.DataFrame) -> None:
    """Render a compact dashboard for normalized finance transactions."""
    kpis = st.columns(3)
    kpis[0].metric("Total income", f"{df['amount_income'].sum():,.2f}")
    kpis[1].metric("Total expenses", f"{df['amount_expense'].sum():,.2f}")
    kpis[2].metric("Net cash flow", f"{df['signed_amount'].sum():,.2f}")

    chart_df = df.copy()
    chart_df["month_sort"] = pd.to_datetime(chart_df["date"], errors="coerce").dt.strftime("%Y-%m")
    chart_df["month_label"] = pd.to_datetime(chart_df["date"], errors="coerce").dt.strftime("%B")

    monthly = (
        chart_df.groupby(["month_sort", "month_label"], dropna=False)[["amount_income", "amount_expense"]]
        .sum()
        .reset_index()
        .sort_values("month_sort")
    )
    monthly_long = monthly.melt(
        id_vars=["month_sort", "month_label"],
        value_vars=["amount_income", "amount_expense"],
        var_name="Type",
        value_name="Amount",
    )
    monthly_long["Type"] = monthly_long["Type"].replace({"amount_income": "Income", "amount_expense": "Expenses"})
    monthly_long["Amount"] = monthly_long["Amount"].abs()

    st.plotly_chart(
        px.bar(
            monthly_long,
            x="month_label",
            y="Amount",
            color="Type",
            barmode="group",
            title="Monthly Income vs Expenses",
            labels={"month_label": "Month"},
        ),
        use_container_width=True,
    )

    for label_column, title in [
        ("category", "Income vs Expenses by Category"),
        ("project_name", "Income vs Expenses by Project Name"),
    ]:
        comparison = income_expense_comparison(chart_df, label_column)
        st.plotly_chart(
            px.bar(
                comparison,
                x="amount",
                y="label",
                color="Type",
                barmode="group",
                orientation="h",
                title=title,
            ),
            use_container_width=True,
        )


def income_expense_comparison(df: pd.DataFrame, label_column: str) -> pd.DataFrame:
    income = (
        df.groupby(label_column, dropna=False)["amount_income"]
        .sum()
        .abs()
        .reset_index(name="amount")
        .assign(Type="Income")
    )
    expenses = (
        df.groupby(label_column, dropna=False)["amount_expense"]
        .sum()
        .abs()
        .reset_index(name="amount")
        .assign(Type="Expenses")
    )
    return pd.concat([income, expenses], ignore_index=True).rename(columns={label_column: "label"})
