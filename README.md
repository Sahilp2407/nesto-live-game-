# Nesto Cricket Live Platform 🏆🏏

A production-grade, real-time cricket scoring and broadcast platform built for **Nesto Sports Fest Season 4**.

## 🚀 Features

- **Real-time Scoring**: Admin console for ball-by-ball updates.
- **Cinematic Dashboard**: Premium, IPL-style broadcast viewer for users.
- **WebSocket Sync**: Instant updates across all screens using Socket.io.
- **Innings Management**: Professional handling of multi-innings matches, team swaps, and target calculations.
- **Data Export**: One-click Excel download for match statistics (Batting, Bowling, Summary).
- **Cinematic Overlays**: GSAP-driven high-impact animations for 4s, 6s, and Wickets.

## 🛠️ Technology Stack

- **Frontend**: HTML5, Vanilla CSS, GSAP, Socket.io Client.
- **Backend**: Node.js, Express.
- **Real-time**: Socket.io.
- **Reporting**: ExcelJS.

## 🏁 Getting Started

1. **Install Dependencies**:
   ```bash
   npm install
   ```

2. **Start the Server**:
   ```bash
   node server.js
   ```

3. **Access the Panels**:
   - **Admin Console**: `http://localhost:3000/cricket-admin-pro.html`
   - **User Dashboard**: `http://localhost:3000/cricket-dashboard.html`

## 📊 Match Lifecycle

1. **Setup**: Enter team names and match overs.
2. **First Innings**: Score ball-by-ball, manage strikers and bowlers.
3. **Shift Batting**: End the 1st innings to swap teams and set the target.
4. **Second Innings**: Real-time "Runs Needed" and "RRR" calculations.
5. **Match End**: Final report generation and Excel export.

---
Built with ❤️ for Nesto Sports Fest.
