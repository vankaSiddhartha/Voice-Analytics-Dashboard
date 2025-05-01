# 🎙️ Voice-Analytics-Dashboard

**Voice-Analytics-Dashboard** is a fun prototype that allows users to control and view Excel-based charts using voice commands. Upload your `.xls` file and simply **speak** commands to navigate through different data visualizations.

🔗 **Live Demo**: [voice-ruby-zeta.vercel.app](https://voice-ruby-zeta.vercel.app/)  

---

## ⚙️ Features

- 🗣️ **Voice Command Navigation**
  - Say commands like:
    - “Show bar graph”
    - “Display pie chart of sales”
    - “Next chart”
  - Commands trigger chart updates based on the uploaded Excel sheet

- 📊 **Excel Sheet Upload**
  - Upload `.xls` file with tabular data
  - Automatically parses and visualizes with charts (bar, pie, line)

- 🌐 **Built with Next.js + React**
  - Simple, frontend-only prototype
  - Uses the Web Speech API for real-time speech-to-text

- 🧪 **Prototype Level**
  - Made for learning & fun!
  - Limited voice command set and basic parsing logic

---

## 🛠️ Tech Stack

- **Next.js** – React framework
- **React** – UI development
- **xlsx** – For reading Excel files
- **Chart.js / Recharts** – Chart rendering
- **Web Speech API** – Voice-to-text input

---

## 🚀 Getting Started Locally

```bash
git clone https://github.com/vankaSiddhartha/Voice-Analytics-Dashboard.git
cd Voice-Analytics-Dashboard
npm install
npm run dev
