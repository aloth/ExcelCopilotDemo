# Mastering Data Analysis with Excel Copilot

[![YouTube](https://img.shields.io/badge/YouTube-Excel%20Copilot%20Demo-red?logo=youtube&style=flat)](https://www.youtube.com/watch?v=1Ug1htXEww4)
[![Blog](https://img.shields.io/badge/Blog-alexloth.com-blue?style=flat)](https://alexloth.com/how-to-use-copilot-in-excel-mastering-excel-with-m365-copilot-data-analysis-tutorial/)
[![Follow on X](https://img.shields.io/twitter/follow/xlth?style=social)](https://x.com/xlth)

> Hands-on dataset + copy-paste prompts for the **Excel Copilot Data Analysis Demo**  

### Quick links
- **Video (watch)**: https://www.youtube.com/watch?v=1Ug1htXEww4  
- **Blog post**: https://alexloth.com/how-to-use-copilot-in-excel-mastering-excel-with-m365-copilot-data-analysis-tutorial/  
- **Download dataset**: [`Donors.xlsx`](https://github.com/aloth/ExcelCopilotDemo/raw/refs/heads/main/Donors.xlsx)

---

## Quickstart (5 minutes)
1. **Download** `Donors.xlsx` from this repo.  
2. **Upload** the file to **OneDrive** (or a SharePoint document library) — Copilot requires cloud storage.  
3. **Open** the workbook in Excel (desktop or Excel for web) and ensure **AutoSave** is ON.  
4. Select any cell inside the dataset and press **Ctrl+T** (or *Home → Format as Table*) to ensure the range is an official **Excel Table**. The table name should be visible in the **Table Design** tab (e.g., `Donors`).  
5. Open the **Copilot** pane in Excel and paste one of the example prompts from the section below (e.g., *Days since last donation*). Review the suggested formula/changes and click **Insert**.

---

## Prerequisites Checklist

Use the following table to quickly verify that your environment is ready for Copilot.

| Requirement | Description | How to Verify |
| :--- | :--- | :--- |
| **Base License** | A qualifying Microsoft 365 subscription (e.g., Business Standard/Premium, E3/E5). | Check your subscription status in your Microsoft 365 account portal. |
| **Copilot License** | The specific "Copilot for Microsoft 365" add-on license. | Confirm with your M365 administrator or check your assigned licenses. |
| **Application Version** | Microsoft 365 Apps on the Current Channel or Monthly Enterprise Channel. | In Excel, go to `File > Account`. Your channel is listed under "About Excel". |
| **File Location** | The `.xlsx` or `.xlsm` file must be saved in OneDrive or SharePoint. | Check that the AutoSave toggle in the top-left of Excel is "On". |
| **Data Format** | Your data must be formatted as an official Excel Table. | Select your data. If the "Table Design" tab appears in the ribbon, it's a table. |

---

## Copy-paste prompts — ready to use

### 1\. Visual Enhancement

  * **Objective:** To visually distinguish high-value donors directly within the dataset for quick identification.
  * **The Prompt:**
    ```plaintext
    Add a colorful icon if Donation\_Amount is more than 100.
    ```
  * **Deconstructing the Prompt:**
      * **Goal:** "Add a colorful icon." This is a clear, action-oriented instruction.
      * **Source:** Copilot infers the source is the active Excel Table.
      * **Context/Condition:** "...if Donation\_Amount is more than 100." This provides the specific logical condition for the action.
      * **Expectation:** The output should be a visual change (an icon) rather than a new data value.
  * **Copilot's Expected Output & Action:** Copilot will propose applying a conditional formatting rule to the `Donation_Amount` column. A preview of the change will be shown in the Copilot pane. The user must review this suggestion and click the "Insert" or "Apply" button to commit the change to the worksheet.
  * **Deeper Insight & Pro-Tips:** Copilot leverages Excel's native conditional formatting capabilities. This means users can request more sophisticated visual rules. For example, try prompts like: "Apply a red-yellow-green color scale to the Donation\_Amount column" or "Highlight the top 10% of donations in green". This demonstrates that Copilot is a natural language interface for existing powerful Excel features.

### 2\. Days Since Last Donation (Recency Analysis)

  * **Objective:** To calculate the recency of each donor's last contribution, a key metric for identifying active versus lapsed donors.
  * **The Prompt:**
    ```plaintext
    Calculate the number of days since each donor’s last donation and add this as a new column.
    ```
  * **Deconstructing the Prompt:**
      * **Goal:** "Calculate the number of days..." and "add this as a new column." This clearly defines both the calculation and the desired output format.
      * **Source:** Copilot will identify the column containing donation dates (e.g., `Last_Donation_Date`) and use the current date as a reference for the calculation.
      * **Context:** The phrase "since each donor's last donation" provides the business context for the calculation.
      * **Expectation:** A new column will be added to the table containing the calculated number of days.
  * **Copilot's Expected Output & Action:** Copilot will suggest a new column, likely named `Days Since Last Donation`, and provide the Excel formula it will use (e.g., `=TODAY()-`). After reviewing the suggested formula and column, the user clicks **Insert Column**.
  * **Deeper Insight & Pro-Tips:** This calculation is a cornerstone of **RFM (Recency, Frequency, Monetary) analysis**, a powerful marketing technique for customer segmentation. After creating this "Recency" column, a powerful follow-up prompt would be: "Group donors into segments based on 'Days Since Last Donation': 0-30 days as 'Active', 31-90 as 'At Risk', and \>90 as 'Lapsed' in a new column called 'Donor Status'." This connects a simple calculation to a high-value strategic analysis.

### 3\. Campaign Effectiveness

  * **Objective:** To analyze the performance of different communication channels by comparing their impact on donation amounts.
  * **The Prompt:**
    ```plaintext
    Analyze the effectiveness of different communication methods (email, phone) on donation amounts. Identify which method is most effective for different donor segments.
    ```
  * **Deconstructing the Prompt:**
      * **Goal:** "Analyze the effectiveness..." and "Identify which method is most effective..." This is a more open-ended analytical request.
      * **Source:** Copilot will need to use columns related to communication methods (e.g., `Campaign_Type`) and donation amounts (`Donation_Amount`).
      * **Context:** The prompt specifies the comparison between "email" and "phone" and asks for analysis across "donor segments."
      * **Expectation:** The output will likely be a summary, a PivotTable, or a chart that aggregates donation data by campaign type.
  * **Copilot's Expected Output & Action:** Copilot will likely generate a PivotTable or chart that shows the average or total donation amount for each communication method. It may also provide a brief textual summary of its findings. The user can then request to add this analysis to a new sheet in the workbook.
  * **Deeper Insight & Pro-Tips:** This prompt demonstrates Copilot's ability to perform light analytical work. To get more granular results, the user could refine the prompt: "Create a PivotTable showing the average 'Donation\_Amount' by 'Campaign\_Type' and 'Occupation'." This forces a more specific and structured output, giving the user greater control over the analysis.

### 4\. Occupation Impact

  * **Objective:** To understand how a donor's profession correlates with their giving behavior.
  * **The Prompt:**
    ```plaintext
    Examine how a donor’s occupation affects their donation behavior, including donation amount and frequency. Identify which occupations are associated with the highest donations.
    ```
  * **Deconstructing the Prompt:**
      * **Goal:** "Examine how a donor's occupation affects their donation behavior..." and "Identify which occupations are associated with the highest donations."
      * **Source:** This requires the `Occupation`, `Donation_Amount`, and potentially a donation count or frequency metric.
      * **Context:** The prompt asks for a multi-faceted analysis, considering both amount and frequency.
      * **Expectation:** A summary, chart, or PivotTable grouping data by occupation.
  * **Copilot's Expected Output & Action:** Copilot will likely generate an analysis, such as a bar chart showing the total or average donation amount per occupation, and provide a summary of the top-donating professions. This can be added to a new sheet for further review.
  * **Deeper Insight & Pro-Tips:** While Copilot can identify correlations, it cannot infer causation. The analysis might show that "Engineers" donate more, but it doesn't explain why. The user's domain expertise is crucial for interpreting these results. A good follow-up prompt for deeper analysis would be: "Show me the distribution of 'Campaign\_Type' for the top 3 occupations." This could reveal if certain professions are more responsive to specific outreach methods.

### 5\. Lifetime Value (LTV)

  * **Objective:** To calculate the total historical value of each donor to the organization.
  * **The Prompt:**
    ```plaintext
    Compute the lifetime value (total donations over time) for each donor and add it as a new column.
    ```
  * **Deconstructing the Prompt:**
      * **Goal:** "Compute the lifetime value..." and "add it as a new column."
      * **Source:** This requires identifying unique donors (e.g., by `Donor_ID`) and summing their `Donation_Amount`.
      * **Context:** The term "lifetime value" provides clear business context.
      * **Expectation:** A new column in the table showing the aggregated donation total for each unique donor.
  * **Copilot's Expected Output & Action:** This is a more complex task that may require Copilot to generate a multi-step formula, perhaps involving `SUMIFS` or other aggregation functions. It will present the formula and the new column for review before the user commits the change by clicking **Insert Column**.
  * **Deeper Insight & Pro-Tips:** Calculating LTV is a fundamental task in customer/donor relationship management. Once this column is created, it becomes a powerful variable for further analysis. A user could then ask, "Show me the average 'Lifetime Value' by 'Occupation'," combining this new metric with existing data for richer insights.

### 6\. Seasonal Donation Trends

  * **Objective:** To identify patterns in donation activity across different seasons of the year.
  * **The Prompt:**
    ```plaintext
    Identify seasonal trends by calculating the total donations for each donor during different seasons (spring, summer, fall, winter) and add these as new columns.
    ```
  * **Deconstructing the Prompt:**
      * **Goal:** "Identify seasonal trends by calculating the total donations..." and "add these as new columns."
      * **Source:** This requires the `Donation_Amount` and `Last_Donation_Date` columns.
      * **Context:** The prompt explicitly defines the seasons to be used.
      * **Expectation:** Four new columns (e.g., `Spring_Donations`, `Summer_Donations`, etc.) added to the table, with each column containing the sum of donations for that season.
  * **Copilot's Expected Output & Action:** Copilot will need to generate complex formulas that can categorize dates into seasons and then sum the corresponding donations. It will present these formulas and the new columns for user approval. The user will then click **Insert Columns**.
  * **Deeper Insight & Pro-Tips:** This analysis helps inform the timing of fundraising campaigns. After generating these columns, a user could ask Copilot to visualize the results: "Create a stacked bar chart showing the total seasonal donations for the top 5 occupations." This would provide a powerful visual for strategic planning, highlighting which donor segments are most active during which seasons.

---

## Advanced: Python, forecasting & ML
Copilot can generate Python code that runs inside Excel (if Excel+Python is enabled). Example advanced prompts (use only after copying the file to OneDrive and ensuring Python in Excel is available):

```plaintext
Perform k-means clustering (k=3) on donors using 'Days Since Last Donation', 'Donation\_Amount', and 'Lifetime\_Value'. Show cluster centers and add a new column 'ClusterID'.
```

Or:

```plaintext
Train a linear regression model to predict Donation\_Amount using Days Since Last Donation and Lifetime\_Value. Provide coefficients, R² and a scatter plot of predictions vs actuals.
```

---

## Best practices

- **Always verify** Copilot-generated formulas and logic. Trust, but verify.  
- **Use exact table/column names** in prompts to avoid ambiguity.  
- **Privacy**: avoid uploading sensitive personal data to public repos. Anonymize donor PII when sharing.  
- **Storage**: Copilot only works with `.xlsx`/`.xlsm` stored in OneDrive/SharePoint. Local files will not work.  
- **Scalability**: Copilot works with large tables (up to millions of cells) but very complex transforms on huge tables may be slow.

---

## Contribute

Contributions welcome! Suggested contribution flow: 
Fork → create branch → PR with clear description.  

---

## License

This repository is licensed under **GNU GPL v3.0** — see `LICENSE` for details.

---

## Further Reading

Thank you for joining this data exploration. To continue your journey and deepen your understanding of AI and data strategy:

  * To dive deeper into the strategic impact of AI on business decision-making, explore my book **[Decisively Digital](https://alexloth.com/decisively-digital/)**.
  * For a comprehensive guide to data visualization and business intelligence principles that complement your work in Excel, check out my book **[Teach Yourself VISUALLY Power BI](https://alexloth.com/power-bi-book/)**.
  * For more articles, tutorials, and insights, visit my blog at **[alexloth.com](https://alexloth.com)**.
