const express = require('express');
const http = require('http');
const { Server } = require('socket.io');
const path = require('path');
const cors = require('cors');
const fs = require('fs');
const ExcelJS = require('exceljs');

const firebase = require('firebase/compat/app');
require('firebase/compat/firestore');

const app = express();
const server = http.createServer(app);
const io = new Server(server, {
    cors: { origin: "*", methods: ["GET", "POST"] },
    maxHttpBufferSize: 1e8 // 100 MB
});

app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ limit: '10mb', extended: true }));
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

// --- FIREBASE SERVER INIT ---
const firebaseConfig = {
    apiKey: "AIzaSyB6vcdBeao5TinoXaumw49ZNk38sj-gL6w",
    authDomain: "nesto-cricket.firebaseapp.com",
    databaseURL: "https://nesto-cricket-default-rtdb.firebaseio.com",
    projectId: "nesto-cricket",
    storageBucket: "nesto-cricket.firebasestorage.app",
    messagingSenderId: "942124875188",
    appId: "1:942124875188:web:91276dd24dc9e83393c6e1"
};
firebase.initializeApp(firebaseConfig);
const db = firebase.firestore();
const matchRef = db.collection('matches').doc('live_match');

const loadState = () => {
    // 1. Initial Load from Local File (Sync fallback)
    try {
        if (fs.existsSync(STATE_FILE)) {
            const data = fs.readFileSync(STATE_FILE, 'utf8');
            matchState = JSON.parse(data);
            console.log('Local state loaded');
        }
    } catch (e) { console.error('Local load error:', e); }

    // 2. Real-time Cloud Sync (Primary)
    matchRef.onSnapshot((doc) => {
        if (doc.exists) {
            console.log("☁️  Cloud State Received");
            matchState = doc.data();
            // Don't call saveState() here to avoid loop, but emit to sockets
            io.emit('matchUpdate', matchState);
        }
    }, err => console.error("Firestore listen error:", err));
}

const saveState = () => {
    fs.writeFileSync(STATE_FILE, JSON.stringify(matchState, null, 2));
    io.emit('matchUpdate', matchState);
    matchRef.set(matchState).catch(err => console.error("Firestore push error:", err));
};

// Start initialization
loadState();

// --- API ROUTES ---

// Get current live state
app.get('/api/match/live', (req, res) => {
    res.json(matchState);
});

// Update entire state (Legacy support, now auto-pruned)
app.post('/api/match/update', (req, res) => {
    let newState = req.body;
    if (newState.innings) {
        newState.innings.forEach(inn => {
            if (inn.history) {
                inn.history = inn.history.slice(-5).map(h => { const c = {...h}; delete c.history; return c; });
            }
        });
    }
    matchState = { ...matchState, ...newState };
    saveState();
    res.json({ success: true, state: matchState });
});

app.post('/api/run', (req, res) => {
    try {
        const { runs, isBoundary, strikerId, bowlerId } = req.body;
        const inn = matchState.innings[matchState.currentInningsIdx];
        if (!inn) return res.status(400).json({ error: "No active innings" });

        // Update Total
        inn.totalRuns += parseInt(runs) || 0;

        // Update Batsman
        const sId = strikerId || inn.strikerId;
        if (sId && inn.batsmen[sId]) {
            const b = inn.batsmen[sId];
            b.runs += parseInt(runs) || 0;
            b.balls++;
            if (isBoundary === 4) b.fours++;
            if (isBoundary === 6) b.sixes++;
        }
        
        // Update Bowler
        const bId = bowlerId || inn.currentBowlerId;
        if (bId && inn.bowlers[bId]) {
            const bowler = inn.bowlers[bId];
            bowler.runs += parseInt(runs) || 0;
            bowler.balls++;
        }

        // Tracking Over
        inn.totalBalls++;
        inn.currentOver = Math.floor(inn.totalBalls / 6);
        inn.currentBall = inn.totalBalls % 6;

        // Swap strike if odd runs
        if (runs % 2 !== 0) {
            const temp = inn.strikerId;
            inn.strikerId = inn.nonStrikerId;
            inn.nonStrikerId = temp;
        }

        matchState.lastEvent = (runs === 4) ? 'four' : (runs === 6 ? 'six' : 'run');
        matchState.lastEventId++;
        
        saveState();
        res.json({ success: true, state: matchState });
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

app.post('/api/wicket', (req, res) => {
    try {
        const { playerOutId, type, bowlerId } = req.body;
        const inn = matchState.innings[matchState.currentInningsIdx];
        if (!inn) return res.status(400).json({ error: "No active innings" });

        inn.totalWickets++;
        inn.totalBalls++; // Wicket is a ball
        inn.currentOver = Math.floor(inn.totalBalls / 6);
        inn.currentBall = inn.totalBalls % 6;

        if (playerOutId && inn.batsmen[playerOutId]) {
            const p = inn.batsmen[playerOutId];
            p.status = 'out';
            p.dismissal = type || 'Wicket';
            p.balls++;
        }

        const bId = bowlerId || inn.currentBowlerId;
        if (bId && inn.bowlers[bId]) {
            const bowler = inn.bowlers[bId];
            bowler.wickets++;
            bowler.balls++;
        }

        matchState.lastEvent = 'wicket';
        matchState.lastEventId++;
        saveState();
        res.json({ success: true, state: matchState });
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

app.post('/api/ball', (req, res) => {
    try {
        const inn = matchState.innings[matchState.currentInningsIdx];
        if (!inn) return res.status(400).json({ error: "No active innings" });

        inn.totalBalls++;
        inn.currentOver = Math.floor(inn.totalBalls / 6);
        inn.currentBall = inn.totalBalls % 6;

        if (inn.strikerId && inn.batsmen[inn.strikerId]) inn.batsmen[inn.strikerId].balls++;
        if (inn.currentBowlerId && inn.bowlers[inn.currentBowlerId]) inn.bowlers[inn.currentBowlerId].balls++;

        saveState();
        res.json({ success: true, state: matchState });
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});
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
    socket.emit('matchUpdate', matchState);

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
