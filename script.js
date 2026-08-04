// =========================================
// LEADERBOARD + RESULTS
// =========================================

// =========================================
// KONFIGURATION
// =========================================

const excelFile =
    "scorekarte_leaderboard.xlsx";

const STORAGE_KEY =
    "golfPlayedScores";

// =========================================
// SUPABASE
// =========================================

const SUPABASE_URL =
    "https://pdwgmbheawyulnxvzzzb.supabase.co";

const SUPABASE_KEY =
    "sb_publishable_Fw5AztwZqtUJfQMFJTWvvA_aqH-83hG";

let supabaseClient = null;


// =========================================
// SUPABASE VERBINDUNG
// =========================================

function initializeSupabase() {

    if (
        typeof supabase === "undefined"
    ) {

        console.error(
            "Supabase-Bibliothek wurde nicht geladen."
        );

        return false;

    }


    supabaseClient =
        supabase.createClient(
            SUPABASE_URL,
            SUPABASE_KEY
        );


    console.log(
        "Supabase-Verbindung hergestellt."
    );


    return true;

}

// Startzeilen der vier Scorekarten

const roundStartRows = {

    1: 1,
    2: 34,
    3: 67,
    4: 100

};


// =========================================
// EXCEL-SPALTEN
// =========================================
//
// PAR:
// Lenz         = D
// Schlehhuber  = M
// Raser        = V
//
// SpV:
// Lenz         = F
// Schlehhuber  = O
// Raser        = X
//
// Die Spaltennummern beginnen bei 1.
// =========================================

const playerColumns = {

    Lenz: {

        par: 4,
        spv: 6

    },

    Schlehhuber: {

        par: 13,
        spv: 15

    },

    Raser: {

        par: 22,
        spv: 24

    }

};


// =========================================
// SPIELER
// =========================================

const players = [

    {

        id: "Lenz",
        name: "Ludwig LENZ",
        country: "GER"

    },

    {

        id: "Schlehhuber",
        name: "Andreas SCHLEHHUBER",
        country: "GER"

    },

    {

        id: "Raser",
        name: "Peter RASER",
        country: "AUS"

    }

];


// =========================================
// DATEN
// =========================================

let workbook = null;

let holeData = {};

let playedScores = {};

let golfClubNames = {};

let currentRound = 1;

let currentThrough = 0;


// =========================================
// EXCEL LADEN
// =========================================

async function loadExcel() {

    try {

        const response =
            await fetch(excelFile);


        if (!response.ok) {

            throw new Error(
                "Excel-Datei konnte nicht geladen werden."
            );

        }


        const arrayBuffer =
            await response.arrayBuffer();


        workbook =
            XLSX.read(
                arrayBuffer,
                {
                    type: "array"
                }
            );


        readExcelData();


        console.log(
            "Excel erfolgreich geladen."
        );


       await updateAll();

    } catch (error) {

        console.error(
            "Fehler beim Laden der Excel-Datei:",
            error
        );

    }

}


// =========================================
// EXCEL-DATEN AUSLESEN
// =========================================
//
// Pro Loch lesen wir:
//
// PAR
// SpV für jeden Spieler
//
// Die Schläge kommen ausschließlich
// aus der Eingabe-App.
// =========================================

function readExcelData() {

    const sheet =
        workbook.Sheets["3-Spieler"];


    if (!sheet) {

        throw new Error(
            'Das Tabellenblatt "3-Spieler" wurde nicht gefunden.'
        );

    }


    // =========================================
    // GOLFPLATZ AUS EXCEL
    // =========================================

    const golfClubCells = {

        1: "B1",
        2: "B34",
        3: "B67",
        4: "B100"

    };


    for (
        const round of [1, 2, 3, 4]
    ) {

        const cell =
            sheet[
                golfClubCells[round]
            ];


        golfClubNames[round] =
            cell && cell.v
                ? String(cell.v)
                : "";

    }


    // =========================================
    // RUNDEN UND LÖCHER
    // =========================================

    for (
        let round = 1;
        round <= 4;
        round++
    ) {


        for (
            let hole = 1;
            hole <= 18;
            hole++
        ) {


            const startRow =
                roundStartRows[round];


            const excelRow =
                startRow +
                8 +
                hole;


            const key =
                `${round}_${hole}`;


            holeData[key] = {

                round: round,
                hole: hole,
                par: null,
                spv: {}

            };


            const parValues = [];


            // =================================
            // PAR UND SPV
            // =================================

            for (
                const player
                of Object.keys(playerColumns)
            ) {

                const columns =
                    playerColumns[player];


                // -----------------------------
                // PAR
                // -----------------------------

                const parCell =
                    sheet[
                        XLSX.utils.encode_cell({

                            r:
                                excelRow - 1,

                            c:
                                columns.par - 1

                        })
                    ];


                const par =
                    getCellNumber(
                        parCell
                    );


                if (
                    par !== null
                ) {

                    parValues.push(
                        par
                    );

                }


                // -----------------------------
                // SPV
                // -----------------------------

                const spvCell =
                    sheet[
                        XLSX.utils.encode_cell({

                            r:
                                excelRow - 1,

                            c:
                                columns.spv - 1

                        })
                    ];


                holeData[key].spv[player] =
                    getCellNumber(
                        spvCell
                    );

            }


            if (
                parValues.length > 0
            ) {

                holeData[key].par =
                    parValues[0];

            }

        }

    }


    console.log(
        "PAR- und SpV-Daten:",
        holeData
    );


    console.log(
        "Golfplätze:",
        golfClubNames
    );

}


// =========================================
// EXCEL-ZELLE ALS ZAHL
// =========================================

function getCellNumber(cell) {

    if (!cell) {

        return null;

    }


    const value =
        cell.v;


    if (
        value === undefined ||
        value === null ||
        value === ""
    ) {

        return null;

    }


    const number =
        Number(value);


    if (
        !Number.isFinite(number)
    ) {

        return null;

    }


    return number;

}


// =========================================
// EINGABEN LADEN
// =========================================

function loadPlayedScores() {

    const saved =
        localStorage.getItem(
            STORAGE_KEY
        );


    if (!saved) {

        return {};

    }


    try {

        const data =
            JSON.parse(saved);


        if (
            typeof data !== "object" ||
            data === null
        ) {

            return {};

        }


        return data;

    } catch (error) {

        console.error(
            "Gespeicherte Schläge konnten nicht gelesen werden.",
            error
        );


        return {};

    }

}

// =========================================
// SCORES AUS SUPABASE LADEN
// =========================================

async function loadScoresFromSupabase() {

    if (!supabaseClient) {

        console.error(
            "Supabase ist nicht verbunden."
        );

        return {};

    }


    const {
        data,
        error
    } = await supabaseClient
        .from("scores")
        .select(
            "round, hole, player, score"
        );


    if (error) {

        console.error(
            "Scores konnten nicht aus Supabase geladen werden:",
            error
        );

        return {};

    }


    const scores = {};


    data.forEach(
        row => {

            const key =
                `${row.round}_${row.hole}_${row.player}`;


            scores[key] =
                Number(row.score);

        }
    );


    return scores;

}

// =========================================
// SCORE HOLEN
// =========================================

function getScore(
    round,
    hole,
    playerId
) {

    const key =
        `${round}_${hole}_${playerId}`;


    if (
        playedScores[key] === undefined
    ) {

        return null;

    }


    const score =
        Number(
            playedScores[key]
        );


    if (
        !Number.isFinite(score)
    ) {

        return null;

    }


    return score;

}


// =========================================
// PAR HOLEN
// =========================================

function getPar(
    round,
    hole
) {

    const key =
        `${round}_${hole}`;


    if (
        !holeData[key]
    ) {

        return null;

    }


    return holeData[key].par;

}


// =========================================
// SPV HOLEN
// =========================================

function getSpv(
    round,
    hole,
    playerId
) {

    const key =
        `${round}_${hole}`;


    if (
        !holeData[key]
    ) {

        return null;

    }


    const value =
        holeData[key].spv[playerId];


    if (
        value === undefined ||
        value === null
    ) {

        return null;

    }


    return Number(value);

}


// =========================================
// BRUTTO-STABLEFORD
// =========================================

function calculateGrossStableford(
    score,
    par
) {

    if (
        score === null ||
        par === null
    ) {

        return 0;

    }


    const points =
        2 -
        (
            score -
            par
        );


    return Math.max(
        0,
        points
    );

}


// =========================================
// NETTO-STABLEFORD
// =========================================

function calculateNetStableford(
    score,
    par,
    spv
) {

    if (
        score === null ||
        par === null ||
        spv === null
    ) {

        return 0;

    }


    const netPar =
        par +
        spv;


    const points =
        2 -
        (
            score -
            netPar
        );


    return Math.max(
        0,
        points
    );

}


// =========================================
// LWS
// =========================================

function calculateLwsForHole(
    scores
) {

    if (
        scores.length !== 3
    ) {

        return null;

    }


    if (
        scores.some(
            score =>
                score === null
        )
    ) {

        return null;

    }


    const lowest =
        Math.min(
            ...scores
        );


    const winners =
        scores.filter(
            score =>
                score === lowest
        ).length;


    if (
        winners === 1
    ) {

        return scores.map(
            score =>
                score === lowest
                    ? 2
                    : 0
        );

    }


    return scores.map(
        score =>
            score === lowest
                ? 1
                : 0
    );

}


// =========================================
// SPIELERWERTE ZURÜCKSETZEN
// =========================================

function resetPlayerTotals() {

    players.forEach(
        player => {

            player.totalShots = 0;

            player.totalPar = 0;

            player.scoreToPar = 0;

            player.totalGrossStableford = 0;

            player.totalNetStableford = 0;

            player.totalLws = 0;

            player.playedHoles = 0;

        }
    );

}


// =========================================
// LEADERBOARD BERECHNEN
// =========================================

function calculateLeaderboardData() {

    resetPlayerTotals();


    for (
        let round = 1;
        round <= 4;
        round++
    ) {


        for (
            let hole = 1;
            hole <= 18;
            hole++
        ) {


            const par =
                getPar(
                    round,
                    hole
                );


            if (
                par === null
            ) {

                continue;

            }


            const scores =
                players.map(
                    player =>
                        getScore(
                            round,
                            hole,
                            player.id
                        )
                );


            const allPlayersPlayed =
                scores.every(
                    score =>
                        score !== null
                );


            if (
                !allPlayersPlayed
            ) {

                continue;

            }


            players.forEach(
                (
                    player,
                    index
                ) => {


                    const score =
                        scores[index];


                    const spv =
                        getSpv(
                            round,
                            hole,
                            player.id
                        );


                    player.totalShots +=
                        score;


                    player.totalPar +=
                        par;


                    player.playedHoles++;


                    player.totalGrossStableford +=
                        calculateGrossStableford(
                            score,
                            par
                        );


                    player.totalNetStableford +=
                        calculateNetStableford(
                            score,
                            par,
                            spv
                        );

                }
            );


            const lws =
                calculateLwsForHole(
                    scores
                );


            if (lws) {

                players.forEach(
                    (
                        player,
                        index
                    ) => {

                        player.totalLws +=
                            lws[index];

                    }
                );

            }

        }

    }


    players.forEach(
        player => {

            if (
                player.playedHoles === 0
            ) {

                player.scoreToPar =
                    0;

                return;

            }


            player.scoreToPar =
                player.totalShots -
                player.totalPar;

        }
    );

    // =========================================
// NEUE GESAMTWERTUNG
// =========================================
//
// Vier gleich gewichtete Wertungen:
//
// 1. Brutto zu PAR  → niedrigster Wert gewinnt
// 2. LWS            → höchster Wert gewinnt
// 3. Netto STB      → höchster Wert gewinnt
// 4. Brutto STB     → höchster Wert gewinnt
//
// Führende bzw. geteilt führende Spieler
// erhalten jeweils 1 Wertungspunkt.
//

players.forEach(
    player => {

        player.rankingPoints = 0;

    }
);


// -----------------------------------------
// 1. BRUTTO ZU PAR
// -----------------------------------------

const playersWithScores =
    players.filter(
        player =>
            player.playedHoles > 0
    );


if (playersWithScores.length > 0) {

    const bestScoreToPar =
        Math.min(
            ...playersWithScores.map(
                player =>
                    player.scoreToPar
            )
        );


    playersWithScores.forEach(
        player => {

            if (
                player.scoreToPar ===
                bestScoreToPar
            ) {

                player.rankingPoints++;

            }

        }
    );

}


// -----------------------------------------
// 2. LWS
// -----------------------------------------

if (playersWithScores.length > 0) {

    const bestLws =
        Math.max(
            ...playersWithScores.map(
                player =>
                    player.totalLws
            )
        );


    playersWithScores.forEach(
        player => {

            if (
                player.totalLws ===
                bestLws
            ) {

                player.rankingPoints++;

            }

        }
    );

}


// -----------------------------------------
// 3. NETTO STABLEFORD
// -----------------------------------------

if (playersWithScores.length > 0) {

    const bestNetStableford =
        Math.max(
            ...playersWithScores.map(
                player =>
                    player.totalNetStableford
            )
        );


    playersWithScores.forEach(
        player => {

            if (
                player.totalNetStableford ===
                bestNetStableford
            ) {

                player.rankingPoints++;

            }

        }
    );

}


// -----------------------------------------
// 4. BRUTTO STABLEFORD
// -----------------------------------------

if (playersWithScores.length > 0) {

    const bestGrossStableford =
        Math.max(
            ...playersWithScores.map(
                player =>
                    player.totalGrossStableford
            )
        );


    playersWithScores.forEach(
        player => {

            if (
                player.totalGrossStableford ===
                bestGrossStableford
            ) {

                player.rankingPoints++;

            }

        }
    );

}

   players.sort(
    (
        a,
        b
    ) => {

        // =================================
        // 1. GESAMT-WERTUNGSPUNKTE
        // Mehr ist besser.
        // =================================

        if (
            a.rankingPoints !==
            b.rankingPoints
        ) {

            return (
                b.rankingPoints -
                a.rankingPoints
            );

        }


        // =================================
        // 2. BRUTTO ZU PAR
        // Weniger ist besser.
        // =================================

        if (
            a.scoreToPar !==
            b.scoreToPar
        ) {

            return (
                a.scoreToPar -
                b.scoreToPar
            );

        }


        // =================================
        // 3. LWS
        // Mehr ist besser.
        // =================================

        if (
            a.totalLws !==
            b.totalLws
        ) {

            return (
                b.totalLws -
                a.totalLws
            );

        }


        // =================================
        // 4. NACHNAME
        // Alphabetisch.
        // =================================

        const lastNameA =
            a.name
                .trim()
                .split(/\s+/)
                .pop()
                .toLowerCase();


        const lastNameB =
            b.name
                .trim()
                .split(/\s+/)
                .pop()
                .toLowerCase();


        return lastNameA.localeCompare(
            lastNameB,
            "de"
        );

    }
);

}


// =========================================
// AKTUELLEN TURNIERSTAND ERMITTELN
// =========================================

function updateRoundInfo() {

    let lastRound = 1;

    let lastHole = 0;


    for (
        let round = 1;
        round <= 4;
        round++
    ) {


        for (
            let hole = 1;
            hole <= 18;
            hole++
        ) {


            const scores =
                players.map(
                    player =>
                        getScore(
                            round,
                            hole,
                            player.id
                        )
                );


            const complete =
                scores.every(
                    score =>
                        score !== null
                );


            if (
                complete
            ) {

                lastRound =
                    round;

                lastHole =
                    hole;

            }

        }

    }


    currentRound =
        lastRound;


    currentThrough =
        lastHole;


    const roundInfo =
        document.getElementById(
            "round-info"
        );


    if (!roundInfo) {

        return;

    }


    const golfClub =
        golfClubNames[currentRound]
            || "GC LOIPERSDORF";


    if (
        currentThrough > 0
    ) {

        roundInfo.textContent =
            `ROUND ${currentRound} | ${golfClub} | THR ${currentThrough}`;

    } else {

        roundInfo.textContent =
            `ROUND ${currentRound} | ${golfClub}`;

    }

}


// =========================================
// SCORE FORMATIEREN
// =========================================

function formatScoreToPar(
    score,
    playedHoles
) {

    if (
        playedHoles === 0
    ) {

        return "–";

    }


    if (
        score === 0
    ) {

        return "E";

    }


    if (
        score > 0
    ) {

        return `+${score}`;

    }


    return `${score}`;

}


// =========================================
// LEADERBOARD LEEREN
// =========================================

function clearLeaderboard() {

    const leaderboard =
        document.getElementById(
            "leaderboard"
        );


    if (!leaderboard) {

        console.error(
            "Element #leaderboard wurde nicht gefunden."
        );


        return null;

    }


    leaderboard.innerHTML = "";


    return leaderboard;

}


// =========================================
// LEADERBOARD DARSTELLEN
// =========================================

function renderLeaderboard() {

    const leaderboard =
        clearLeaderboard();


    if (!leaderboard) {

        return;

    }


    players.forEach(
        (
            player,
            index
        ) => {

            const playerRow =
                document.createElement(
                    "article"
                );


            playerRow.className =
                "player-row";


            playerRow.innerHTML = `

                <div class="rank">
                    ${index + 1}
                </div>


                <div class="player-info">

                    <div class="player-name">
                        ${player.name}
                    </div>

                    <div class="player-country">
                        ${player.country}
                    </div>

                </div>


                <div class="score-over-par">

                    ${formatScoreToPar(
                        player.scoreToPar,
                        player.playedHoles
                    )}

                </div>


                <div class="stats">


                    <div class="stat">

                        <span class="stat-label">
                            NETTO
                        </span>

                        <span class="stat-value">

                            ${
                                player.playedHoles > 0
                                    ? player.totalNetStableford
                                    : "–"
                            }

                        </span>

                    </div>


                    <div class="stat">

                        <span class="stat-label">
                            BRUTTO
                        </span>

                        <span class="stat-value">

                            ${
                                player.playedHoles > 0
                                    ? player.totalGrossStableford
                                    : "–"
                            }

                        </span>

                    </div>


                    <div class="stat">

                        <span class="stat-label">
                            LWS
                        </span>

                        <span class="stat-value">

                            ${
                                player.playedHoles > 0
                                    ? player.totalLws
                                    : "–"
                            }

                        </span>

                    </div>


                </div>

            `;


            leaderboard.appendChild(
                playerRow
            );

        }
    );

}


// =========================================
// SCORECARD: SCORE-STYLING
// =========================================
//
// Par = normale Zahl
// Birdie = 1 Kreis
// Eagle oder besser = 2 Kreise
// Bogey = 1 Quadrat
// Double Bogey oder schlechter = 2 Quadrate
// =========================================

function getScoreClass(
    score,
    par
) {

    if (
        score === null ||
        par === null
    ) {

        return "";

    }


    const difference =
        score - par;


    if (
        difference <= -2
    ) {

        return "circle-2";

    }


    if (
        difference === -1
    ) {

        return "circle-1";

    }


    if (
        difference === 1
    ) {

        return "square-1";

    }


    if (
        difference >= 2
    ) {

        return "square-2";

    }


    return "";

}


// =========================================
// SCORECARD SCORE DARSTELLEN
// =========================================

function renderScoreValue(
    score,
    par
) {

    if (
        score === null
    ) {

        return "";

    }


    const className =
        getScoreClass(
            score,
            par
        );


    if (
        className === ""
    ) {

        return `
            <span class="score-value">
                ${score}
            </span>
        `;

    }


    return `
        <span class="score-value ${className}">
            ${score}
        </span>
    `;

}


// =========================================
// SCORECARD ERSTELLEN
// =========================================

function renderPlayerScorecard(
    player,
    round
) {

    let frontPar = 0;
    let backPar = 0;

    let frontScore = 0;
    let backScore = 0;

    let frontGross = 0;
    let backGross = 0;

    let frontNet = 0;
    let backNet = 0;


    // =====================================
    // SCORECARD-HTML
    // =====================================

    let html = `

        <article class="result-player-card">

            <div class="result-player-name">
                ${player.name}
            </div>

            <div class="scorecard-scroll">

                <table class="scorecard">

                    <thead>

                        <tr>

                            <th class="row-label">
                                HOLE
                            </th>
    `;


    // Löcher 1–9

    for (
        let hole = 1;
        hole <= 9;
        hole++
    ) {

        html += `
            <th>
                ${hole}
            </th>
        `;

    }


    html += `
            <th class="section-total">
                OUT
            </th>
    `;


    // Löcher 10–18

    for (
        let hole = 10;
        hole <= 18;
        hole++
    ) {

        html += `
            <th>
                ${hole}
            </th>
        `;

    }


    html += `

            <th class="section-total">
                IN
            </th>

            <th class="total-column">
                TOT
            </th>

        </tr>

    </thead>

    <tbody>

        <tr class="par-row">

            <td class="row-label">
                PAR
            </td>
    `;


    // PAR 1–9

    for (
        let hole = 1;
        hole <= 9;
        hole++
    ) {

        const par =
            getPar(
                round,
                hole
            );


        if (par !== null) {

            frontPar += par;

        }


        html += `
            <td>
                ${
                    par !== null
                        ? par
                        : ""
                }
            </td>
        `;

    }


    html += `
            <td class="section-total">
                ${frontPar}
            </td>
    `;


    // PAR 10–18

    for (
        let hole = 10;
        hole <= 18;
        hole++
    ) {

        const par =
            getPar(
                round,
                hole
            );


        if (par !== null) {

            backPar += par;

        }


        html += `
            <td>
                ${
                    par !== null
                        ? par
                        : ""
                }
            </td>
        `;

    }


    const totalPar =
        frontPar +
        backPar;


    html += `

            <td class="section-total">
                ${backPar}
            </td>

            <td class="total-column">
                ${totalPar}
            </td>

        </tr>


        <tr>

            <td class="row-label">
                SCORE
            </td>
    `;


    // SCORE 1–9

    for (
        let hole = 1;
        hole <= 9;
        hole++
    ) {

        const score =
            getScore(
                round,
                hole,
                player.id
            );


        const par =
            getPar(
                round,
                hole
            );


        if (score !== null) {

            frontScore += score;

            frontGross +=
                calculateGrossStableford(
                    score,
                    par
                );

            frontNet +=
                calculateNetStableford(
                    score,
                    par,
                    getSpv(
                        round,
                        hole,
                        player.id
                    )
                );

        }


        html += `
            <td>
                ${renderScoreValue(
                    score,
                    par
                )}
            </td>
        `;

    }


    html += `
            <td class="section-total">
                ${
                    frontScore > 0
                        ? frontScore
                        : ""
                }
            </td>
    `;


    // SCORE 10–18

    for (
        let hole = 10;
        hole <= 18;
        hole++
    ) {

        const score =
            getScore(
                round,
                hole,
                player.id
            );


        const par =
            getPar(
                round,
                hole
            );


        if (score !== null) {

            backScore += score;

            backGross +=
                calculateGrossStableford(
                    score,
                    par
                );

            backNet +=
                calculateNetStableford(
                    score,
                    par,
                    getSpv(
                        round,
                        hole,
                        player.id
                    )
                );

        }


        html += `
            <td>
                ${renderScoreValue(
                    score,
                    par
                )}
            </td>
        `;

    }


    const totalScore =
        frontScore +
        backScore;


    html += `

            <td class="section-total">
                ${
                    backScore > 0
                        ? backScore
                        : ""
                }
            </td>

            <td class="total-column">
                ${
                    totalScore > 0
                        ? totalScore
                        : ""
                }
            </td>

        </tr>


        <tr>

            <td class="row-label">
                STB
            </td>
    `;


    // STABLEFORD 1–9

    for (
        let hole = 1;
        hole <= 9;
        hole++
    ) {

        const score =
            getScore(
                round,
                hole,
                player.id
            );


        const par =
            getPar(
                round,
                hole
            );


        const points =
            calculateGrossStableford(
                score,
                par
            );


        html += `
            <td>
                ${
                    score !== null
                        ? points
                        : ""
                }
            </td>
        `;

    }


    html += `
            <td class="section-total">
                ${
                    frontScore > 0
                        ? frontGross
                        : ""
                }
            </td>
    `;


    // STABLEFORD 10–18

    for (
        let hole = 10;
        hole <= 18;
        hole++
    ) {

        const score =
            getScore(
                round,
                hole,
                player.id
            );


        const par =
            getPar(
                round,
                hole
            );


        const points =
            calculateGrossStableford(
                score,
                par
            );


        html += `
            <td>
                ${
                    score !== null
                        ? points
                        : ""
                }
            </td>
        `;

    }


    const totalGross =
        frontGross +
        backGross;


    html += `

            <td class="section-total">
                ${
                    backScore > 0
                        ? backGross
                        : ""
                }
            </td>

            <td class="total-column">
                ${
                    totalGross > 0
                        ? totalGross
                        : ""
                }
            </td>

        </tr>

    </tbody>

</table>

</div>


<div class="result-summary">

    <div class="result-summary-item">

        <span class="result-summary-label">
            PAR
        </span>

        <span class="result-summary-value">
            ${totalPar}
        </span>

    </div>


    <div class="result-summary-item">

        <span class="result-summary-label">
            BRUTTO
        </span>

        <span class="result-summary-value">
            ${
                totalScore > 0
                    ? totalScore
                    : "–"
            }
        </span>

    </div>


    <div class="result-summary-item">

        <span class="result-summary-label">
            BRUTTO STB
        </span>

        <span class="result-summary-value">
            ${
                totalGross > 0
                    ? totalGross
                    : "–"
            }
        </span>

    </div>

</div>

</article>

`;


    return html;

}


// =========================================
// RESULTS DARSTELLEN
// =========================================

function renderResults() {

    const container =
        document.getElementById(
            "results-container"
        );


    if (!container) {

        return;

    }


    const select =
        document.getElementById(
            "results-round"
        );


    if (!select) {

        return;

    }


    const round =
        Number(
            select.value
        );


    const course =
        document.getElementById(
            "results-course"
        );


    if (course) {

        const golfClub =
            golfClubNames[round]
                || "GC LOIPERSDORF";


        course.textContent =
            `ROUND ${round} | ${golfClub}`;

    }


    container.innerHTML = "";


    // Alphabetische Reihenfolge
    const sortedPlayers =
        [...players].sort(
            (
                a,
                b
            ) =>
                a.name.localeCompare(
                    b.name
                )
        );


    sortedPlayers.forEach(
        player => {

            container.innerHTML +=
                renderPlayerScorecard(
                    player,
                    round
                );

        }
    );

}


// =========================================
// GESAMT AKTUALISIEREN
// =========================================

async function updateAll() {

    playedScores =
        await loadScoresFromSupabase();


    if (
        !workbook
    ) {

        return;

    }


    calculateLeaderboardData();

    renderLeaderboard();

    updateRoundInfo();

    renderResults();

}


// =========================================
// SEITENNAVIGATION
// =========================================

// =========================================
// SEITENNAVIGATION
// =========================================

function setupNavigation() {

    const navLeaderboard =
        document.getElementById(
            "nav-leaderboard"
        );


    const navResults =
        document.getElementById(
            "nav-results"
        );


    const navNews =
        document.getElementById(
            "nav-news"
        );


    const leaderboardSection =
        document.getElementById(
            "leaderboard-section"
        );


    const resultsSection =
        document.getElementById(
            "results-section"
        );


    const newsSection =
        document.getElementById(
            "news-section"
        );


    if (
        !navLeaderboard ||
        !navResults ||
        !navNews ||
        !leaderboardSection ||
        !resultsSection ||
        !newsSection
    ) {

        console.error(
            "Navigationselemente wurden nicht vollständig gefunden."
        );

        return;

    }


    function showPage(
        page
    ) {

        leaderboardSection.classList.add(
            "hidden"
        );

        resultsSection.classList.add(
            "hidden"
        );

        newsSection.classList.add(
            "hidden"
        );


        navLeaderboard.classList.remove(
            "active"
        );

        navResults.classList.remove(
            "active"
        );

        navNews.classList.remove(
            "active"
        );


        if (
            page === "leaderboard"
        ) {

            leaderboardSection.classList.remove(
                "hidden"
            );

            navLeaderboard.classList.add(
                "active"
            );

        }


        if (
            page === "results"
        ) {

            resultsSection.classList.remove(
                "hidden"
            );

            navResults.classList.add(
                "active"
            );

            renderResults();

        }


        if (
    page === "news"
) {

    newsSection.classList.remove(
        "hidden"
    );

    navNews.classList.add(
        "active"
    );

    loadNews();

}

    } // <-- DIESE KLAMMER HAT IN DEINER VERSION GEFEHLT


    navLeaderboard.addEventListener(
        "click",
        () => {

            showPage(
                "leaderboard"
            );

        }
    );


    navResults.addEventListener(
        "click",
        () => {

            showPage(
                "results"
            );

        }
    );


    navNews.addEventListener(
        "click",
        () => {

            showPage(
                "news"
            );

        }
    );


    const resultsRound =
        document.getElementById(
            "results-round"
        );


    if (
        resultsRound
    ) {

        resultsRound.addEventListener(
            "change",
            () => {

                renderResults();

            }
        );

    }

}


// =========================================
// SHEETJS LADEN
// =========================================

function loadSheetJS() {

    return new Promise(
        (
            resolve,
            reject
        ) => {


            if (
                typeof XLSX !== "undefined"
            ) {

                resolve();

                return;

            }


            const script =
                document.createElement(
                    "script"
                );


            script.src =
                "https://cdn.jsdelivr.net/npm/xlsx/dist/xlsx.full.min.js";


            script.onload =
                resolve;


            script.onerror =
                () => reject(
                    new Error(
                        "SheetJS konnte nicht geladen werden."
                    )
                );


            document.head.appendChild(
                script
            );

        }
    );

}

// =========================================
// NEWS LADEN
// =========================================

async function loadNews() {

    const container =
        document.getElementById(
            "news-container"
        );

    if (!container) {
        return;
    }

    if (!supabaseClient) {

        console.error(
            "Supabase ist nicht verbunden."
        );

        return;
    }


    // =========================================
    // NEWS LADEN
    // =========================================

    const {
        data: news,
        error: newsError
    } = await supabaseClient
        .from("news")
        .select(
            "id, created_at, title, content, published"
        )
        .eq(
            "published",
            true
        )
        .order(
            "created_at",
            {
                ascending: false
            }
        );


    if (newsError) {

        console.error(
            "News konnten nicht geladen werden:",
            newsError
        );

        container.innerHTML = `
            <div class="news-empty">
                News konnten nicht geladen werden.
            </div>
        `;

        return;
    }


    container.innerHTML = "";


    if (
        !news ||
        news.length === 0
    ) {

        container.innerHTML = `
            <div class="news-empty">
                Noch keine News vorhanden.
            </div>
        `;

        return;
    }


    // =========================================
    // JEDE NEWS-MELDUNG DARSTELLEN
    // =========================================

    for (
        const post of news
    ) {

        // -----------------------------------------
        // MEDIA ZUR NEWS LADEN
        // -----------------------------------------

        const {
            data: media,
            error: mediaError
        } = await supabaseClient
            .from("news_media")
            .select(
                "id, news_id, storage_path, media_type, created_at"
            )
            .eq(
                "news_id",
                post.id
            )
            .order(
                "created_at",
                {
                    ascending: true
                }
            );


        if (mediaError) {

            console.error(
                "News-Media konnte nicht geladen werden:",
                mediaError
            );

        }


        // =========================================
// ARTIKEL ERSTELLEN
// =========================================

const article =
    document.createElement(
        "article"
    );

article.className =
    "news-item";


// =========================================
// DATUM
// =========================================

const date =
    new Date(
        post.created_at
    );


const formattedDate =
    date.toLocaleDateString(
        "de-DE",
        {
            day: "2-digit",
            month: "short",
            year: "numeric"
        }
    ).toUpperCase();


// =========================================
// GRUNDSTRUKTUR
// =========================================

article.innerHTML = `

    <div class="news-date">
        ${formattedDate}
    </div>

    <div class="news-content">

        <div class="news-text-content">

            <h2 class="news-title">
                ${escapeHtml(post.title)}
            </h2>

            <p class="news-text">
                ${escapeHtml(post.content)}
            </p>

        </div>

    </div>

`;
const newsContent =
    article.querySelector(".news-content");
        
// =========================================
// BILD ZUR NEWS
// =========================================



        // =========================================
        // BILDER
        // =========================================

        if (
            media &&
            media.length > 0
        ) {

            const mediaContainer =
                document.createElement(
                    "div"
                );

            mediaContainer.className =
                "news-media";


            media.forEach(
                mediaItem => {

                    // Nur Bilder anzeigen
                    if (
                        mediaItem.media_type !==
                        "image"
                    ) {

                        return;

                    }


                    // ---------------------------------
                    // ÖFFENTLICHE STORAGE-URL
                    // ---------------------------------

                    const {
                        data: publicUrlData
                    } =
                        supabaseClient
                            .storage
                            .from(
                                "news-media"
                            )
                            .getPublicUrl(
                                mediaItem.storage_path
                            );


                    const imageUrl =
                        publicUrlData?.publicUrl;


                    if (!imageUrl) {

                        console.error(
                            "Keine öffentliche URL für:",
                            mediaItem.storage_path
                        );

                        return;

                    }


                    // ---------------------------------
                    // BILD
                    // ---------------------------------

                    const image =
                        document.createElement(
                            "img"
                        );

                    image.className =
                        "news-image";


                    image.src =
                        imageUrl;


                    image.alt =
                        post.title ||
                        "News Bild";


                    image.loading =
                        "lazy";


                    mediaContainer.appendChild(
                        image
                    );

                }
            );


            // Nur hinzufügen, wenn tatsächlich
            // Bilder vorhanden sind.

            if (
                mediaContainer.children.length > 0
            ) {

               newsContent.appendChild(
    mediaContainer
);

            }

        }


        // =========================================
        // ARTIKEL AUF SEITE SETZEN
        // =========================================

        container.appendChild(
            article
        );

    }

}




// =========================================
// HTML-SICHERHEIT
// =========================================

function escapeHtml(value) {

    const div =
        document.createElement(
            "div"
        );

    div.textContent =
        value ?? "";

    return div.innerHTML;

}

// =========================================
// AUTOMATISCHE AKTUALISIERUNG
// =========================================
//
// Alle 5 Sekunden werden neue Eingaben
// übernommen.
// =========================================

setInterval(
    updateAll,
    5000
);



// =========================================
// START
// =========================================

(async function () {

    try {

        initializeSupabase();

        setupNavigation();

        await loadSheetJS();

        await loadExcel();

    } catch (error) {

        console.error(
            error
        );

    }

})();
