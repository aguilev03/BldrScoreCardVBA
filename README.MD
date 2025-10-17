🧮 Builder ScoreCard VBA

A VBA automation tool for Excel that analyzes builder data, counts project statuses, and color-codes performance metrics — all while excluding weekends from aging calculations.

This macro is designed to help track On-Time-In-Full (OTIF) completion, occupancy status, and average project age, producing a formatted monthly summary automatically.

🚀 Features

Counts business days only (Mon–Fri) between start and completion dates.

Automatically colors rows based on project status and aging thresholds.

Calculates:

Total records per month

OTIF counts

Occupied vs. Vacant counts

Average age for each category

Exports a clean summary to a Formatted data sheet.

Fully compatible with Excel’s NETWORKDAYS function for accurate weekday calculations.

📂 Project Structure
BldrScoreCardVBA/
│
├── CountRows_BusinessDays_VBA.txt   # Full VBA source code (copy into a VBA module)
├── README.md                        # This documentation
└── Example.xlsx                     # (optional) sample workbook showing macro in action

⚙️ Setup Instructions

Download or clone the repository:

git clone https://github.com/aguilev03/BldrScoreCardVBA.git


Open your Excel workbook.

Press Alt + F11 to open the VBA editor.

In the menu, click Insert → Module.

Copy and paste the contents of CountRows_BusinessDays_VBA.txt.

Adjust column constants at the top of the script if your sheet layout differs:

Const COL_START As String = "E"
Const COL_COMPLETED As String = "F"


Run the macro:

Press F5 in the VBA editor, or

Create a button in Excel linked to
CountRowsByMonthAndOTIFAndOccupiedAndAverageAgeAndVacantAndColor.

🎨 Color Coding Rules
Status	Age Range	Color
Vacant	0–3 days	🟩 Light Green
Vacant	4–5 days	🟨 Yellow
Vacant	6+ days	🟥 Red
Occupied	0–5 days	🟩 Light Green
Occupied	6–8 days	🟨 Yellow
Occupied	9+ days	🟥 Red
🧠 How It Works

Uses NETWORKDAYS to compute business days (excluding weekends).

Optionally supports a list of holidays (Range("AA2:AA50")) if you wish to skip them too.

Writes the computed "Age" into column G.

Summarizes all monthly metrics in a separate sheet named Formatted data.

💡 Customization

You can easily modify:

Color thresholds (change Select Case logic)

Columns used for Start, Completed, Age, or Status

Add a holiday exclusion list by enabling:

ageBiz = BusinessDays(CDate(startDate), CDate(endDate), wsBuilder.Range("AA2:AA50"))

🧰 Requirements

Microsoft Excel with VBA enabled

Macro permissions enabled in your workbook

Tested on Office 365 and Excel 2021 (Windows)

🤝 Contributing

Pull requests are welcome!
If you’d like to add features (like holiday handling or advanced dashboards), open an issue or submit a PR.

🧾 License

This project is licensed under the MIT License — feel free to use and modify it for your own reporting workflows.

👤 Author

Evan Aguilar
💼 Flooring & Claims Specialist • 🧑‍💻 Tech Hobbyist • 📊 Excel Automation Enthusiast
GitHub: @aguilev03
