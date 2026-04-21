const express = require('express');
const http = require('http');
const { Server } = require('socket.io');
const path = require('path');
const cors = require('cors');
const fs = require('fs');
const ExcelJS = require('exceljs');

const app = express();
const server = http.createServer(app);
const io = new Server(server, {
    cors: { origin: "*", methods: ["GET", "POST"] }
});

app.use(cors());
app.use(express.json()); // Essential for API calls
app.use(express.static(path.join(__dirname)));

const STATE_FILE = path.join(__dirname, 'match_state.json');

// Global Match State (Initialized from file if exists)
let matchState = {
    phase: 'select',
    status: 'not_started',
    selectedTeams: [],
    format: 'T20',
    maxOvers: 20,
    toss: { winner: null, decision: null },
    setup: { striker: null, nonStriker: null, bowler: null },
    currentInningsIdx: 0,
    innings: [],
    lastEvent: null,
    lastEventId: 0
};

// Load state on startup
if (fs.existsSync(STATE_FILE)) {
    try {
        const data = fs.readFileSync(STATE_FILE, 'utf8');
        matchState = JSON.parse(data);
        console.log('Match state loaded from storage');
    } catch (e) {
        console.error('Error loading state:', e);
    }
}

const saveState = () => {
    fs.writeFileSync(STATE_FILE, JSON.stringify(matchState, null, 2));
    io.emit('match-update', matchState);
};

// --- API ROUTES ---

// Get current live state
app.get('/api/match/live', (req, res) => {
    res.json(matchState);
});

// Update entire state
app.post('/api/match/update', (req, res) => {
    matchState = { ...matchState, ...req.body };
    saveState();
    res.json({ success: true, state: matchState });
});

// Start Match API
app.post('/api/match/start', (req, res) => {
    matchState.status = 'live';
    matchState.phase = 'match';
    saveState();
    res.json({ success: true, status: 'live' });
});

// Export Match Data to Excel
app.get('/api/match/export', async (req, res) => {
    try {
        if (!matchState.innings || matchState.innings.length === 0) {
            return res.status(400).send("No match data to export.");
        }

        const workbook = new ExcelJS.Workbook();
        
        // 1. MATCH SUMMARY SHEET
        const summarySheet = workbook.addWorksheet('Match Summary');
        summarySheet.columns = [
            { header: 'Metric', key: 'metric', width: 25 },
            { header: 'Details', key: 'details', width: 35 }
        ];
        
        summarySheet.addRows([
            { metric: 'Team A', details: matchState.selectedTeams[0] || '---' },
            { metric: 'Team B', details: matchState.selectedTeams[1] || '---' },
            { metric: 'Format', details: matchState.format },
            { metric: 'Max Overs', details: matchState.maxOvers },
            { metric: 'Toss Winner', details: matchState.toss.winner || '---' },
            { metric: 'Toss Decision', details: matchState.toss.decision || '---' },
            { metric: 'Status', details: matchState.status.toUpperCase() }
        ]);

        // 2. BATTING SCORECARD SHEET
        const battingSheet = workbook.addWorksheet('Batting Scorecard');
        battingSheet.columns = [
            { header: 'Innings', key: 'innings', width: 10 },
            { header: 'Team', key: 'team', width: 25 },
            { header: 'Player Name', key: 'name', width: 25 },
            { header: 'Runs', key: 'runs', width: 10 },
            { header: 'Balls', key: 'balls', width: 10 },
            { header: 'Strike Rate', key: 'sr', width: 15 },
            { header: 'Status', key: 'status', width: 15 }
        ];

        matchState.innings.forEach((inn, idx) => {
            Object.values(inn.batsmen).forEach(b => {
                battingSheet.addRow({
                    innings: idx + 1,
                    team: inn.battingTeam,
                    name: b.name,
                    runs: b.runs,
                    balls: b.balls,
                    sr: b.balls > 0 ? ((b.runs / b.balls) * 100).toFixed(2) : '0.00',
                    status: b.status.toUpperCase()
                });
            });
        });

        // 3. BOWLING STATS SHEET
        const bowlingSheet = workbook.addWorksheet('Bowling Stats');
        bowlingSheet.columns = [
            { header: 'Innings', key: 'innings', width: 10 },
            { header: 'Team', key: 'team', width: 25 },
            { header: 'Player Name', key: 'name', width: 25 },
            { header: 'Overs', key: 'overs', width: 10 },
            { header: 'Runs', key: 'runs', width: 10 },
            { header: 'Wickets', key: 'wickets', width: 10 },
            { header: 'Economy', key: 'eco', width: 15 }
        ];

        matchState.innings.forEach((inn, idx) => {
            Object.values(inn.bowlers).forEach(b => {
                bowlingSheet.addRow({
                    innings: idx + 1,
                    team: inn.bowlingTeam,
                    name: b.name,
                    overs: `${Math.floor(b.balls / 6)}.${b.balls % 6}`,
                    runs: b.runs,
                    wickets: b.wickets,
                    eco: b.balls > 0 ? (b.runs / (b.balls / 6)).toFixed(2) : '0.00'
                });
            });
        });

        // Style the headers
        [summarySheet, battingSheet, bowlingSheet].forEach(sheet => {
            sheet.getRow(1).font = { bold: true };
            sheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0E0E0' } };
        });

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=Nesto_Match_Stats.xlsx');

        await workbook.xlsx.write(res);
        res.end();

    } catch (e) {
        console.error("Export Error:", e);
        res.status(500).send("Error generating Excel file.");
    }
});

// Generic Reset API
app.post('/api/match/reset', (req, res) => {
    matchState = {
        phase: 'select', status: 'not_started', selectedTeams: [], format: 'T20',
        maxOvers: 20, toss: { winner: null, decision: null },
        setup: { striker: null, nonStriker: null, bowler: null },
        currentInningsIdx: 0, innings: [], lastEvent: null, lastEventId: 0
    };
    saveState();
    res.json({ success: true });
});

// --- SOCKET.IO HANDLING ---

io.on('connection', (socket) => {
    console.log('Client connected:', socket.id);
    socket.emit('match-update', matchState);

    socket.on('admin-update', (state) => {
        matchState = state;
        saveState();
        console.log('Match state synced via Socket');
    });

    socket.on('disconnect', () => {
        console.log('Client disconnected');
    });
});

const PORT = process.env.PORT || 3000;
server.listen(PORT, () => {
    console.log(`\n🚀 NESTO LIVE ENGINE ACTIVE`);
    console.log(`📡 Server Port: ${PORT}`);
    console.log(`🛠️  Admin: http://localhost:${PORT}/cricket-admin-pro.html`);
    console.log(`📊 Viewer: http://localhost:${PORT}/cricket-dashboard.html\n`);
});
