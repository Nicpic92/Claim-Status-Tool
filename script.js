// --- FINAL CONFIGURATION ---
const REQUIRED_HEADERS = ['Claim Number', 'Claim Edits', 'Claim Notes', 'Claim Status', 'Claim State', 'Clean Age', 'Age', 'Payer', 'NetworkStatus', 'DSNP or Non DSNP', 'Received Date', 'DOSFromDate', 'DOSToDate']; 
// UPDATED: Included 'Claim Edits' in PV_WORK_COLUMNS
const PV_WORK_COLUMNS = ['Payer', 'Billing Provider', 'Billing Provider Tax ID', 'Billing Provider NPI', 'ProviderFullName', 'ProviderNPI', 'NetworkStatus', 'Plan Name', 'Claim Edits', 'Claim Notes'];
const DATE_COLUMNS = ['Received Date', 'DOSFromDate', 'DOSToDate']; 
const DATE_FMT = 'mm/dd/yyyy'; 

let globalAllClaims = []; 
let globalHeaders = [];   
let globalGroupingMap = null; 

let pvSubTeamNamesGlobal = [];
let pvSubTeamTotalsGlobal = {};
let pvTotalGlobal = 0;

// --- UPDATED: New Team Structure with NO Default Logic ---
const CLAIMS_FUNCTIONS = {
    CLAIMS: 'Claims Team',
    PV: 'PV Team (Provider Ops)',
    UNASSIGNED: 'Needs Assignment (Initial Review)' 
};

let claimsSubTeamNamesGlobal = [CLAIMS_FUNCTIONS.CLAIMS]; 
let claimsTeamTotalGlobal = 0;

const AGING_BUCKETS = ['0-20 Queue', '21-27 Priority', '28-30 Critical', '31+ Backlog'];

// **PERSISTENCE FIX: Load dynamic reassignment rules from localStorage for automatic persistence**
let globalReassignmentRules = JSON.parse(localStorage.getItem('assignmentRules')) || {}; 

// --- PERMANENT RULES OBJECT (Original Default Logic) ---
const PERMANENT_RULES = {};


// --- AGING BUCKET LOGIC ---
function getAgeBucket(age) {
    age = parseInt(age);
    if (isNaN(age)) return 'Age N/A';
    if (age >= 31) return '31+ Backlog';
    if (age >= 28) return '28-30 Critical';
    if (age >= 21) return '21-27 Priority';
    return '0-20 Queue';
}

// --- PV SUB-TEAM LOGIC (for Tab Segregation) ---
function getPvSubTeam(edits, notes) {
    const editsUpper = edits.toUpperCase();
    const notesUpper = notes.toUpperCase();
    
    // --- NEW PV SUB-TEAM CATEGORIZATION ---

    // 1. Provider/Vendor Creation
    if (notesUpper.includes('NEEDS RENDERING AND PAY-TO PROVIDER ADDED') ||
        notesUpper.includes('NEEDS VENDOR ADDED') ||
        notesUpper.includes('NEEDS VENDOR AND RENDERING ADDED') ||
        notesUpper.includes('NEEDS RENDERING ADDED') ||
        notesUpper.includes('NO RENDERING ADDED') ||
        editsUpper.includes('RENDERING PROVIDER DOES NOT EXIST IN PV') ||
        notesUpper.includes('RENDER PHY NEEDS TO BE ADDED') ||
        notesUpper.includes('REQUESTED CLAIM CAN NOT BE MOVED TO PREBATCH WITH NON-VALIDATED VENDOR INFO')
    ) {
        return 'Provider/Vendor Creation';
    }

    // 2. W9/Validation/COB
    if (notesUpper.includes('MISSING W9. REQUESTED') ||
        notesUpper.includes('PROVIDER NOT VALIDATED') ||
        notesUpper.includes('VENDORS ZIP CODE DOESN\'T MATCH') ||
        notesUpper.includes('COB ON FILE')
    ) {
        return 'W9/Validation/COB';
    }

    // 3. Contract/Network Issues
    if (notesUpper.includes('PROVIDER DOESN\'T HAVE A CONTRACT FOR DOS SUBMITTED') ||
        editsUpper.includes('NO ACTIVE CONTRACTS FOUND FOR THIS DOS') ||
        notesUpper.includes('ERROR: NO ACTIVE CONTRACTS FOUND FOR THIS DOS') ||
        editsUpper.includes('NO MATCHING CONTRACT FOUND') ||
        notesUpper.includes('MULTIPLE NETWORK AFFILIATIONS ARE IDENTIFIED') ||
        notesUpper.includes('AUTH_CHECK_OUT OF NETWORK PROVIDER')
    ) {
        return 'Contract/Network Issues';
    }

    // 4. Pay-to Provider Issues
    if (notesUpper.includes('PAY TO PROVIDER DETAILS DOES NOT MATCH WITH THE CONTRACT') ||
        notesUpper.includes('ERROR: PAY TO PROVIDER DETAILS DOES NOT MATCH WITH THE CONTRACT')
    ) {
        return 'Pay-to Provider Issues';
    }

    // 5. Pricing/PBP/Other
    if (editsUpper.includes('PBP NOT FOUND FOR MEMBER') ||
        notesUpper.includes('PART B EMERGENCY CLAIM TO BE PRICED AT LINE LEVEL') ||
        notesUpper.includes('PART B OUTPATIENT CLAIM TO BE PRICED AT LINE LEVEL') ||
        notesUpper.includes('PART A INPATIENT CLAIM TO BE PRICED AT CLAIM LEVEL') ||
        notesUpper.includes('PART B INPATIENT CLAIM TO BE PRICED AT CLAIM LEVEL') ||
        notesUpper.includes('VALIDATE THE ADJUDICATORS INSTRUCTIONS') ||
        notesUpper.includes('NO BENEFIT RULE HITS') ||
        notesUpper.includes('STILL ON HOLD. CONTRACT REQUESTED') ||
        notesUpper.includes('TMS CLAIMS MOVED TO ON HOLD WITH CPT')
    ) {
        return 'Pricing/PBP/Other';
    }

    // 6. Default Fallback (Uncategorized)
    return 'PV Team (Uncategorized)'; 
}

// --- CORE ASSIGNMENT LOGIC (ELIMINATED DEFAULT LOGIC) ---
function assignTeam(edits, notes, customRules) { 
    // This key is always UPPERCASE and is the source of truth for the rule lookup.
    const key = `${sanitizeEdit(edits)}|${sanitizeNote(notes)}`.toUpperCase();
    
    // 1. Check Custom Reassignment Rules (Highest Priority)
    const assignedTeam = customRules && customRules[key];
    
    if (assignedTeam && (assignedTeam === CLAIMS_FUNCTIONS.CLAIMS || assignedTeam === CLAIMS_FUNCTIONS.PV)) {
        return assignedTeam;
    }

    // 2. If no valid custom rule exists, force manual assignment.
    return CLAIMS_FUNCTIONS.UNASSIGNED;
}

// --- Local helper to convert JS Date to Excel Serial Number (Reinstated Fix) ---
function dateToNum(v) {
    if (!(v instanceof Date)) v = new Date(v);
    const D1900 = 25569; 
    const MS_PER_DAY = 86400000;
    let v_num = v.getTime() / MS_PER_DAY + D1900;
    if (v_num < 61) return v_num;
    return v_num + 1; 
}

// --- XLSX Utility for Date Formatting (Reinstated Fix) ---
function formatAsDateCell(dateString) {
    if (!dateString || String(dateString).trim() === '' || String(dateString).toUpperCase() === 'N/A') {
        return { v: dateString }; 
    }
    let dateObj = new Date(dateString);
    if (isNaN(dateObj.getTime())) {
        return { v: dateString }; 
    }
    const v = dateToNum(dateObj); 
    // Ensure type 'n' (number) and format 'z' are set
    return { v: v, t: 'n', z: DATE_FMT };
}

// --- Utility to Clean Headers for Pivot Tables (Aggressive Sanitization) ---
function cleanHeaders(headers) {
    const seen = new Set();
    const cleanedHeaders = headers
        .filter(h => h && String(h).trim() !== '') 
        .map((h, index) => {
            let cleaned = String(h).trim();
            
            // 1. Aggressively sanitize: Keep only alphanumeric and underscores.
            cleaned = cleaned.replace(/[^a-zA-Z0-9_]/g, ''); 
            
            // 2. Ensure a valid name (not empty after cleaning)
            if (!cleaned) {
                cleaned = 'UnnamedCol';
            }

            // 3. Ensure uniqueness
            let unique = cleaned;
            let counter = 1;
            while (seen.has(unique)) {
                unique = cleaned + '_' + counter++;
            }
            seen.add(unique);
            return unique;
        });
    return cleanedHeaders;
}


// --- PDF DOWNLOAD FUNCTION ---
function downloadPDF() {
    const element = document.getElementById('results');
    document.body.classList.add('generating-pdf');

    const opt = {
        margin:       0.3, // Slightly reduced margin for more space
        filename:     'Claims_Analysis_Report.pdf',
        image:        { type: 'jpeg', quality: 0.98 },
        html2canvas:  { scale: 2, useCORS: true },
        jsPDF:        { unit: 'in', format: 'letter', orientation: 'portrait' }
    };

    html2pdf().set(opt).from(element).save().then(() => {
         document.body.classList.remove('generating-pdf');
    }, (err) => {
         document.body.classList.remove('generating-pdf');
         console.error("PDF generation error:", err);
         alert("Error generating PDF. See console.");
    });
}

// --- Utility Functions (File Handling) ---
function showErrorMessage(message) {
    document.getElementById('error-message').innerText = message;
}

// --- NEW FUNCTION: Clear Rules ---
function clearAssignmentRules() {
    if (!confirm("Are you sure you want to clear ALL custom reassignment rules? This cannot be undone and will reset assignments to 'Needs Assignment' for all claims.")) {
        return;
    }
    
    globalReassignmentRules = {};
    localStorage.removeItem('assignmentRules'); 

    // If data is loaded, re-run the assignment with empty rules
    if (globalAllClaims && globalAllClaims.length > 0) {
        applyAssignmentsAndRedraw(false); // Re-run analysis, skip modal interaction
        alert("All custom rules cleared. Reports re-ran. All previous rules-based assignments are now 'Needs Assignment (Initial Review)' unless you reassign them.");
    } else {
        updateReassignmentStatusDisplay(); // Just update the status
        alert("All custom rules cleared. Upload a file to see the impact.");
    }
}
// END NEW FUNCTION

// --- JSON Import/Export Functions (Reinstated) ---
function exportAssignmentRules() {
    if (Object.keys(globalReassignmentRules).length === 0) {
        alert("No custom reassignment rules to export. Run an analysis and apply changes first.");
        return;
    }
    const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(globalReassignmentRules, null, 2));
    const downloadAnchorNode = document.createElement('a');
    downloadAnchorNode.setAttribute("href", dataStr);
    downloadAnchorNode.setAttribute("download", "master_assignment_rules.json");
    document.body.appendChild(downloadAnchorNode);
    downloadAnchorNode.click();
    downloadAnchorNode.remove();
    alert(`Successfully exported ${Object.keys(globalReassignmentRules).length} custom rule(s) to master_assignment_rules.json. This JSON now contains all rules applied in the current analysis.`);
}

function importAssignmentRules(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const importedRules = JSON.parse(e.target.result);
            if (typeof importedRules === 'object' && importedRules !== null) {
                
                // --- FIX: Merge imported rules with existing rules ---
                // Existing rules are overridden by imported rules in case of a key clash.
                globalReassignmentRules = { ...globalReassignmentRules, ...importedRules };
                
                const rulesCount = Object.keys(globalReassignmentRules).length;

                // **PERSISTENCE FIX: Save merged rules to localStorage to persist across reloads**
                localStorage.setItem('assignmentRules', JSON.stringify(globalReassignmentRules)); 
                
                if (!globalAllClaims || globalAllClaims.length === 0) {
                     document.getElementById('reassignment-status').innerText = `Rules loaded: ${rulesCount} custom rule(s) waiting for file upload.`;
                     document.getElementById('reassignment-status').style.color = 'orange';
                     event.target.value = ''; // Clear file input
                     return;
                }
                document.getElementById('reassignmentModal').style.display='none'; // Close modal if open
                
                // Re-run the core analysis logic with the new rules
                applyAssignmentsAndRedraw(false); // Pass false to skip modal interaction
                alert(`Successfully imported and merged ${Object.keys(importedRules).length} rule(s). Total rules: ${rulesCount}. Reports re-ran.`);
            } else {
                showErrorMessage("Import Error: Invalid JSON structure for assignment rules.");
            }
        } catch (e) {
            showErrorMessage("Import Error: Could not parse JSON file. " + e.message);
        }
        event.target.value = ''; // Clear file input
    };
    reader.readAsText(file);
}

// --- NEW UNIVERSAL KEY SANITATION FUNCTION ---
function sanitizeNote(text) {
    // 1. Replace newlines/carriage returns with a single space
    // 2. Collapse multiple spaces into a single space
    // 3. Trim leading/trailing whitespace
    return (text || '--- NO CLAIM NOTES ---')
        .replace(/\r?\n|\r/g, ' ')
        .replace(/\s\s+/g, ' ')
        .trim();
}

function sanitizeEdit(text) {
    return (text || '--- NO CLAIM EDITS ---')
        .replace(/\s\s+/g, ' ')
        .trim();
}


// --- NEW FUNCTION: Redraws the assignment editor table based on a filter ---
function redrawAssignmentEditorTable(filterTeam) {
    const editorContent = document.getElementById('editor-content');
    if (!editorContent || !globalGroupingMap) return;
    
    // --- Alphabetical Sort by Edits/Notes ---
    // NOTE: globalGroupingMap is now keyed by UPPERCASE keys
    let sortedGroups = Array.from(globalGroupingMap.values()).sort((a, b) => {
        if (a.edits < b.edits) return -1;
        if (a.edits > b.edits) return 1;
        if (a.notes < b.notes) return -1;
        if (a.notes > b.notes) return 1;
        return 0;
    });
    
    // --- FILTERING LOGIC ---
    if (filterTeam && filterTeam !== 'ALL') {
        // Filter by the consolidated team (group.team)
        sortedGroups = sortedGroups.filter(group => group.team === filterTeam);
    }
    
    let html = `
        <table class="editor-table">
            <thead>
                <tr>
                    <th style="width: 5%;">Count</th>
                    <th style="width: 20%;">Current Team</th>
                    <th style="width: 30%;">Claim Edits</th>
                    <th style="width: 30%;">Claim Notes</th>
                    <th style="width: 15%;">New Assignment</th>
                </tr>
            </thead>
            <tbody>
    `;
    
    const allTeams = Object.values(CLAIMS_FUNCTIONS);

    sortedGroups.forEach((group) => {
        // The key for rule lookup and data attribute MUST be the UPPERCASE key (which is the key of the group)
        const key = `${group.edits}|${group.notes}`.toUpperCase(); 
        
        // The team currently assigned: This is the result AFTER applying any custom rule (group.team)
        const currentAssignment = group.team; 
        
        // FIX: Prioritize the saved rule for the dropdown selection
        // The value in the dropdown should reflect the saved rule if it exists, otherwise the current (analyzed) team.
        const currentSelectedTeam = globalReassignmentRules[key] || currentAssignment; 
        
        // Build the dropdown options
        let optionsHtml = '';
        allTeams.forEach(team => {
            const isSelected = team === currentSelectedTeam ? 'selected' : '';
            const displayTeamName = team.replace(/\s\(.*\)/m, '');
            
            optionsHtml += `<option value="${team}" ${isSelected}>${displayTeamName}</option>`;
        });
        
        // Determine row color class (based on the currentAssignment, which is the result of the last analysis)
        let rowClass = '';
        if (currentAssignment === CLAIMS_FUNCTIONS.PV) {
            rowClass = 'PV';
        } else if (currentAssignment === CLAIMS_FUNCTIONS.CLAIMS) {
            rowClass = 'CLAIMS';
        } else {
            rowClass = 'UNASSIGNED';
        }


        html += `
            <tr data-key="${key}" class="${rowClass}">
                <td class="count">${group.count.toLocaleString()}</td>
                <td style="font-weight: bold;">${currentAssignment.replace(/\s\(.*\)/m, '')}</td>
                <td class="note-cell">${group.edits}</td>
                <td class="note-cell">${group.notes}</td>
                <td>
                    <select id="select-${key.replace(/[^a-zA-Z0-9]/g, '_')}">
                        ${optionsHtml}
                    </select>
                </td>
            </tr>
        `;
    });

    html += '</tbody></table>';
    editorContent.innerHTML = html;
}

// --- REASSIGNMENT LOGIC (UPDATED WITH FIX) ---
function openReassignmentEditor() {
    if (!globalGroupingMap) {
        showErrorMessage("Please upload and analyze a file first.");
        return;
    }

    const teamFilterSelect = document.getElementById('teamFilterSelect');
    const allTeams = Object.values(CLAIMS_FUNCTIONS);
    
    // 1. Populate the Filter Dropdown (Rebuild every time to ensure counts are current)
    let optionsHtml = '<option value="ALL">--- SHOW ALL TEAMS ---</option>';
    allTeams.forEach(team => {
        // Calculate count based on the **final** applied assignment (group.team)
        const teamCount = Array.from(globalGroupingMap.values()).reduce((sum, group) => {
            return sum + (group.team === team ? group.count : 0);
        }, 0);
        optionsHtml += `<option value="${team}">${team.replace(/\s\(.*\)/m, '')} (${teamCount.toLocaleString()})</option>`;
    });
    teamFilterSelect.innerHTML = optionsHtml;
    
    // Set up the event listener for the dropdown filter change
    teamFilterSelect.onchange = function() {
        redrawAssignmentEditorTable(this.value);
    };

    // 2. Draw the table content using the current filter selection
    redrawAssignmentEditorTable(teamFilterSelect.value);

    // 3. Show the modal
    document.getElementById('reassignmentModal').style.display = 'block';
}

function applyAssignmentsAndRedraw(checkModal = true) {
    if (!globalAllClaims || globalAllClaims.length === 0) {
        alert("Please upload and analyze a file first before applying rules.");
        return;
    }
    
    const modal = document.getElementById('reassignmentModal');
    // FIX: Start with all existing rules to prevent unintentional overwrite
    const newRules = { ...globalReassignmentRules }; 
    let rulesApplied = 0;
    
    if (checkModal && modal.style.display === 'block') {
        // 1. COLLECT ONLY OVERRIDE RULES from the Modal
        document.querySelectorAll('#editor-content tbody tr').forEach(row => {
            const key = row.getAttribute('data-key'); // This is the UPPERCASE key
            const selectElement = row.querySelector('select');
            const selectedTeam = selectElement.value;
            
            // --- CORRECTED RULE SAVING LOGIC ---
            if (selectedTeam === CLAIMS_FUNCTIONS.CLAIMS || selectedTeam === CLAIMS_FUNCTIONS.PV) {
                 // Add or Update the rule (Only CLAIMS or PV can be explicit rules)
                 newRules[key] = selectedTeam;
            } else if (selectedTeam === CLAIMS_FUNCTIONS.UNASSIGNED) {
                 // If set to UNASSIGNED, remove the override rule
                 delete newRules[key];
            }
            // --- END CORRECTED RULE SAVING ---
        });
        
        globalReassignmentRules = newRules; // Update the global map with the merged set
        rulesApplied = Object.keys(globalReassignmentRules).length;

    } else {
        // If checkModal is false (from import or initial load/clear), we just re-apply the existing global rules 
        rulesApplied = Object.keys(globalReassignmentRules).length;
    }
    
    // 2. SAVE RULES (The Persistence Fix) - This is the crucial step
    localStorage.setItem('assignmentRules', JSON.stringify(globalReassignmentRules));
    
    // 3. RE-RUN ASSIGNMENT & REBUILD GROUPING MAP (using the final globalReassignmentRules)
    const newGroupingMap = new Map();
    const activeClaims = globalAllClaims.filter(c => !c.isPrebatch);
    
    globalAllClaims.forEach(claim => {
        const editsRaw = claim['Claim Edits'] || '';
        const notesRaw = claim['Claim Notes'] || '';
        
        const displayEdits = sanitizeEdit(editsRaw);
        const displayNotes = sanitizeNote(notesRaw); 
        const lookupKey = `${displayEdits}|${displayNotes}`.toUpperCase(); // UPPERCASE lookup key

        let newTeam = assignTeam(editsRaw, notesRaw, globalReassignmentRules);
        
        claim['assignedTeam'] = newTeam;
        
        if (newTeam === CLAIMS_FUNCTIONS.PV) {
            claim['pvSubTeam'] = getPvSubTeam(editsRaw, notesRaw);
        } else {
            claim['pvSubTeam'] = null;
        }
        
        // Re-build the grouping map using the CONSISTENT UPPERCASE KEY
        if (!claim.isPrebatch) {
            const groupingKey = lookupKey; 
            
            if (!newGroupingMap.has(groupingKey)) {
                // Store the original case for display but key the map consistently
                newGroupingMap.set(groupingKey, { 
                    count: 0, 
                    edits: displayEdits, // Case-sensitive for display
                    notes: displayNotes, // Case-sensitive for display
                    team: newTeam, 
                    subTeam: claim.pvSubTeam, 
                    sampleClaims: [] 
                });
            }
            const group = newGroupingMap.get(groupingKey);
            group.count++;
            group.team = newTeam; // Update the team property to the newly assigned team for display
        }
    });
    
    globalGroupingMap = newGroupingMap; // Update global map for future editing

    // REINFORCEMENT: Force re-render of the entire results section
    displayResults(newGroupingMap, activeClaims.length);
    
    if (checkModal) modal.style.display = 'none'; 
    
    // REINFORCEMENT: Update status with the final, correct rule count
    document.getElementById('reassignment-status').innerText = `Success: ${rulesApplied} custom rule(s) applied.`;
    document.getElementById('reassignment-status').style.color = rulesApplied > 0 ? 'green' : 'black';
}


function processData() {
    showErrorMessage('');
    globalAllClaims = []; 
    globalHeaders = [];
    globalGroupingMap = null;
    // globalReassignmentRules is NOT cleared here to allow imported/local rules to persist

    const fileInput = document.getElementById('fileInput');
    const textInput = document.getElementById('textInput').value.trim();
    const file = fileInput.files[0];

    if (!file && !textInput) { showErrorMessage("Please upload a file or paste data."); return; }

    if (file) {
        // *** SYNTAX FIX: Corrected split('.pop().toLowerCase()') ***
        const extension = file.name.split('.').pop().toLowerCase();
        const reader = new FileReader();

        reader.onload = function(e) {
            try {
                let data = e.target.result;
                let dataRows;
                if (extension === 'xlsx' || extension === 'xls') {
                    const arr = new Uint8Array(data);
                    const workbook = XLSX.read(arr, { type: 'array' });
                    const sheetName = workbook.SheetNames[0];
                    data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });
                    globalHeaders = data[0].map(h => String(h || '').trim());
                    dataRows = data.slice(1);
                } else { 
                    const lines = data.split(/\r\n|\n/).filter(line => line.trim() !== '');
                    if (lines.length < 2) { showErrorMessage("No Data Found."); return; }
                    const delimiter = lines[0].includes('\t') ? '\t' : (lines[0].includes(',') ? ',' : null);
                    if (!delimiter) { showErrorMessage("Error: Could Not Detect Separator."); return; }
                    globalHeaders = lines[0].split(delimiter).map(h => h.trim());
                    dataRows = lines.slice(1).map(line => line.split(delimiter));
                }

                const hMap = {};
                globalHeaders.forEach((h, i) => hMap[h] = i);
                const indices = REQUIRED_HEADERS.map(h => hMap[h]);

                if (indices.some(i => i === undefined)) { showErrorMessage(`Header Mismatch Error: Missing ${REQUIRED_HEADERS.filter((h, i) => indices[i] === undefined).join(', ')}. Please check your column headers.`); return; }

                // --- Consolidated Grouping and Initial Assignment Logic ---
                const tempGrouping = new Map(); 

                dataRows.forEach(rowData => {
                    if (rowData.length >= globalHeaders.length) {
                        const claim = {};
                        globalHeaders.forEach((h, i) => { claim[h] = String(rowData[i] || '').trim(); });
                        
                        const editsRaw = claim['Claim Edits'] || '';
                        const notesRaw = claim['Claim Notes'] || '';

                        
                        // Determine the CONSOLIDATED KEY for this claim 
                        const displayEdits = sanitizeEdit(editsRaw);
                        const displayNotes = sanitizeNote(notesRaw); 
                        const lookupKey = `${displayEdits}|${displayNotes}`.toUpperCase(); // UPPERCASE lookup key

                        // 1. Calculate Team (Apply custom rules or force UNASSIGNED)
                        let currentTeam = assignTeam(editsRaw, notesRaw, globalReassignmentRules);
                        
                        // Set essential claim properties
                        claim['assignedTeam'] = currentTeam;
                        claim['ageBucket'] = getAgeBucket(claim['Clean Age']);
                        claim['receivedAgeBucket'] = getAgeBucket(claim['Age']);
                        const status = claim['Claim Status'].toUpperCase();
                        const state = claim['Claim State'].toUpperCase();
                        claim['isPrebatch'] = (status === 'PREBATCH' || state === 'PREBATCH' || status === 'DRAFT' || state === 'DRAFT');

                        if (currentTeam === CLAIMS_FUNCTIONS.PV) {
                            claim['pvSubTeam'] = getPvSubTeam(editsRaw, notesRaw);
                        } else {
                            claim['pvSubTeam'] = null;
                        }
                        globalAllClaims.push(claim);

                        // 2. Build the temporary grouping map for consolidation (using the UPPERCASE key)
                        if (!claim.isPrebatch) {
                            const key = lookupKey; 

                            if (!tempGrouping.has(key)) {
                                tempGrouping.set(key, { 
                                    count: 0, 
                                    edits: displayEdits, // Case-sensitive for display
                                    notes: displayNotes, // Case-sensitive for display
                                    teamCounts: {},
                                    sampleClaims: [] 
                                });
                            }
                            const group = tempGrouping.get(key);
                            group.count++;
                            group.teamCounts[currentTeam] = (group.teamCounts[currentTeam] || 0) + 1;
                            if (group.sampleClaims.length < 3) { group.sampleClaims.push(claim['Claim Number']); }
                        }
                    }
                });

                if (globalAllClaims.length === 0) { document.getElementById('results').innerHTML = "<h2>Analysis Complete</h2><p>No valid claims rows were found after processing.</p>"; return; }
                
                // 3. Final Consolidation: Create the official globalGroupingMap (one entry per unique key)
                const finalGroupingMap = new Map();
                tempGrouping.forEach((tempGroup, key) => {
                    // The winning team is determined by the max count, reflecting the actual assignment
                    let winningTeam = CLAIMS_FUNCTIONS.UNASSIGNED;
                    let maxCount = 0;
                    for (const [team, count] of Object.entries(tempGroup.teamCounts)) {
                        if (count > maxCount) {
                            maxCount = count;
                            winningTeam = team;
                        }
                    }
                    
                    // Use the consistent UPPERCASE key
                    finalGroupingMap.set(key, {
                        count: tempGroup.count,
                        edits: tempGroup.edits,
                        notes: tempGroup.notes,
                        team: winningTeam, // The final team for this unique key after rules are applied
                        subTeam: winningTeam === CLAIMS_FUNCTIONS.PV ? getPvSubTeam(tempGroup.edits, tempGroup.notes) : null,
                        sampleClaims: tempGroup.sampleClaims
                    });
                });

                globalGroupingMap = finalGroupingMap; 
                const activeClaims = globalAllClaims.filter(c => !c.isPrebatch);

                displayResults(finalGroupingMap, activeClaims.length);
                document.getElementById('reassignment-section').style.display = 'block';

            } catch (e) {
                showErrorMessage(`Error processing file: ${e.message}`);
            }
        };
        
        if (extension === 'xlsx' || extension === 'xls') {
            reader.readAsArrayBuffer(file);
        } else {
            reader.readAsText(file);
        }
    } else if (textInput) {
         showErrorMessage("Please use the file upload for the final tool.");
    }
}


// --- CORE FUNCTION: Generates the 4 aging tabs and 'All Data' tab for a specific team (Reinstated Fix) ---
function generateTeamAgingTabs(wb, teamClaims, teamName, isFullDataExport = false) {
    // Use globalHeaders for full data export, PV_WORK_COLUMNS headers for regular PV.
    const headersToUse = isFullDataExport ? globalHeaders : PV_WORK_COLUMNS;
    const cleanedHeadersToUse = cleanHeaders(headersToUse);

    // --- 1. Generate Aging Tabs ---
    const agingClaimsGrouped = teamClaims.reduce((acc, claim) => {
        const key = claim.ageBucket; 
        if (!acc[key]) { acc[key] = []; }
        acc[key].push(claim);
        return acc;
    }, {});

    AGING_BUCKETS.forEach(bucket => {
        const claimsInBucket = agingClaimsGrouped[bucket];
        if (!claimsInBucket || claimsInBucket.length === 0) return;

        const sheetName = (isFullDataExport ? 'Full_' : '') + bucket.replace(/[^a-zA-Z0-9]/g, '_').substring(0, 25);
        
        const dataForSheet = claimsInBucket.map(claim => {
            const row = {};
            headersToUse.forEach(originalHeader => {
                row[cleanHeaders([originalHeader])[0]] = claim[originalHeader];
            });
            return row;
        });
        
        const ws = XLSX.utils.json_to_sheet(dataForSheet, { header: cleanedHeadersToUse });
        
        // Apply formatting and AutoFilter
        if (dataForSheet.length > 0 && ws['!ref']) {
            const range = XLSX.utils.decode_range(ws['!ref']);
            ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: range.e }) };

            cleanedHeadersToUse.forEach((header, colIndex) => {
                // Find the original header name from the cleaned one
                const originalHeader = headersToUse.find(h => cleanHeaders([h])[0] === header);
                if (DATE_COLUMNS.includes(originalHeader)) {
                    for (let i = range.s.r + 1; i <= range.e.r; ++i) { 
                        const cellRef = XLSX.utils.encode_cell({c: colIndex, r: i});
                        const cell = ws[cellRef];
                        if (cell && cell.v) {
                            const dateCell = formatAsDateCell(cell.v);
                            if (dateCell.t === 'n') {
                                cell.v = dateCell.v;
                                cell.t = dateCell.t;
                                cell.z = dateCell.z;
                            }
                        }
                    }
                }
            });
        }

        XLSX.utils.book_append_sheet(wb, ws, sheetName);
    });

    // --- 2. Generate All Data Tab ---
    const allDataRows = teamClaims.map(claim => {
        const row = {};
        headersToUse.forEach(originalHeader => {
            row[cleanHeaders([originalHeader])[0]] = claim[originalHeader];
        });
        return row;
    });

    const wsAllData = XLSX.utils.json_to_sheet(allDataRows, { header: cleanedHeadersToUse });
    
    // Apply AutoFilter and formatting to All Data Tab
    if (allDataRows.length > 0 && wsAllData['!ref']) {
        const range = XLSX.utils.decode_range(wsAllData['!ref']);
        wsAllData['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: range.e }) };

        cleanedHeadersToUse.forEach((header, colIndex) => {
            const originalHeader = headersToUse.find(h => cleanHeaders([h])[0] === header);
            if (DATE_COLUMNS.includes(originalHeader)) {
                for (let i = range.s.r + 1; i <= range.e.r; ++i) { 
                    const cellRef = XLSX.utils.encode_cell({c: colIndex, r: i});
                    const cell = wsAllData[cellRef];
                    if (cell && cell.v) {
                        const dateCell = formatAsDateCell(cell.v);
                        if (dateCell.t === 'n') {
                            cell.v = dateCell.v;
                            cell.t = dateCell.t;
                            cell.z = dateCell.z;
                        }
                    }
                }
            }
        });
    }
    
    XLSX.utils.book_append_sheet(wb, wsAllData, (isFullDataExport ? "Full_All_Data" : "All Data"));
}

function downloadClaimsXlsxWorkbook() { 
    if (!globalAllClaims || globalAllClaims.length === 0) {
        alert("No claims data available.");
        return;
    }
    const teamName = CLAIMS_FUNCTIONS.CLAIMS;
    const claims = globalAllClaims.filter(c => c.assignedTeam === teamName && !c.isPrebatch);
    const wb = XLSX.utils.book_new();
    
    // Use generateTeamAgingTabs with isFullDataExport = true to export ALL columns for Claims Team
    generateTeamAgingTabs(wb, claims, teamName, true); 
    
    XLSX.writeFile(wb, `Master_Workbook_Claims.xlsx`);
}

// --- PV FULL DATA WORKBOOK DOWNLOAD FUNCTION (Reinstated) ---
function downloadPvFullDataWorkbook() { 
    if (!globalAllClaims || globalAllClaims.length === 0) {
        alert("No claims data available to create reports.");
        return;
    }
    const teamName = CLAIMS_FUNCTIONS.PV;
    const claims = globalAllClaims.filter(c => c.assignedTeam === teamName && !c.isPrebatch);
    const wb = XLSX.utils.book_new();
    
    // Use generateTeamAgingTabs with isFullDataExport = true to export ALL columns (with PHI)
    generateTeamAgingTabs(wb, claims, teamName, true); 
    
    XLSX.writeFile(wb, `PV_Master_Full_Data_PHI.xlsx`);
}

// --- PV WORKBOOK DOWNLOAD FUNCTION (Stays for non-PHI workload report) ---
function downloadPvXlsxWorkbook() { 
    if (!globalAllClaims || globalAllClaims.length === 0) {
        alert("No claims data available to create reports.");
        return;
    }

    const wb = XLSX.utils.book_new();
    const activeClaims = globalAllClaims.filter(c => !c.isPrebatch); 
    const pvClaims = activeClaims.filter(c => c.assignedTeam === CLAIMS_FUNCTIONS.PV);

    // --- PV SUB-TEAM NAMES UPDATE ---
    pvSubTeamNamesGlobal = [
        'Provider/Vendor Creation',
        'W9/Validation/COB',
        'Contract/Network Issues',
        'Pay-to Provider Issues',
        'Pricing/PBP/Other',
        'PV Team (Uncategorized)'
    ].filter(name => pvSubTeamTotalsGlobal[name] > 0);
    // --- END PV SUB-TEAM NAMES UPDATE ---

    const summaryHeaders = ["PV Sub-Team", "Claim Count", "Percentage of PV Work"];
    const summaryData = [];
    
    pvSubTeamNamesGlobal.forEach(subTeam => {
        const count = pvSubTeamTotalsGlobal[subTeam] || 0;
        const percentage = pvTotalGlobal > 0 ? ((count / pvTotalGlobal) * 100).toFixed(1) + '%' : '0.0%';
        summaryData.push([subTeam, count, percentage]);
    });
    
    const wsSummary = XLSX.utils.aoa_to_sheet([summaryHeaders, ...summaryData]);
    XLSX.utils.book_append_sheet(wb, wsSummary, "PV SUMMARY");
    
    // The headers now include Claim Edits due to the constant update
    const cleanedPvWorkColumns = cleanHeaders(PV_WORK_COLUMNS).concat(['Claim_Count']);

    pvSubTeamNamesGlobal.forEach(subTeam => {
        const claims = pvClaims.filter(c => c.pvSubTeam === subTeam);
        if (claims && claims.length > 0) {
            const deduplicatedMap = {};
            claims.forEach(claim => {
                const key = `${claim['Billing Provider Tax ID']}|${claim['ProviderFullName']}`;
                if (!deduplicatedMap[key]) {
                    deduplicatedMap[key] = { ...claim };
                    deduplicatedMap[key]['Claim Count'] = 1; 
                } else {
                    deduplicatedMap[key]['Claim Count']++;
                }
            });

            const dataForSheet = Object.values(deduplicatedMap).map(item => {
                const row = {};
                PV_WORK_COLUMNS.forEach(originalHeader => {
                    row[cleanHeaders([originalHeader])[0]] = item[originalHeader];
                });
                row['Claim_Count'] = item['Claim Count'];
                return row;
            });
            
            const safeSheetName = subTeam.replace(/[^a-zA-Z0-9]/g, '').substring(0, 30).trim();
            

            const ws = XLSX.utils.json_to_sheet(dataForSheet, { header: cleanedPvWorkColumns });
            
            if (dataForSheet.length > 0 && ws['!ref']) {
                const range = XLSX.utils.decode_range(ws['!ref']);
                ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: range.e }) };
            }

            XLSX.utils.book_append_sheet(wb, ws, safeSheetName);
        }
    });

    XLSX.writeFile(wb, "PV_Team_Workload_Workbook.xlsx");
}

// --- IMPLEMENTATION OF MISSING FUNCTION: downloadReport ---
function downloadReport(teamName, isPvSubTeam = false) {
    if (!globalAllClaims || globalAllClaims.length === 0) {
        alert("No claims data available.");
        return;
    }

    const wb = XLSX.utils.book_new();
    const activeClaims = globalAllClaims.filter(c => !c.isPrebatch);
    
    if (isPvSubTeam) {
        // Logic for a single PV Sub-Team Download (non-PHI, deduplicated)
        const subTeam = teamName;
        const pvClaims = activeClaims.filter(c => c.assignedTeam === CLAIMS_FUNCTIONS.PV);
        const claims = pvClaims.filter(c => c.pvSubTeam === subTeam);
        
        if (claims.length === 0) {
            alert(`No active claims found for PV Sub-Team: ${subTeam}.`);
            return;
        }

        const headersToUse = PV_WORK_COLUMNS;
        const cleanedHeadersToUse = cleanHeaders(headersToUse).concat(['Claim_Count']);

        // Deduplication logic, same as in downloadPvXlsxWorkbook
        const deduplicatedMap = {};
        claims.forEach(claim => {
            const key = `${claim['Billing Provider Tax ID']}|${claim['ProviderFullName']}`;
            if (!deduplicatedMap[key]) {
                deduplicatedMap[key] = { ...claim };
                deduplicatedMap[key]['Claim Count'] = 1; 
            } else {
                deduplicatedMap[key]['Claim Count']++;
            }
        });

        const dataForSheet = Object.values(deduplicatedMap).map(item => {
            const row = {};
            headersToUse.forEach(originalHeader => {
                row[cleanHeaders([originalHeader])[0]] = item[originalHeader];
            });
            row['Claim_Count'] = item['Claim Count'];
            return row;
        });
        
        const safeSheetName = subTeam.replace(/[^a-zA-Z0-9]/g, '').substring(0, 30).trim();
        const ws = XLSX.utils.json_to_sheet(dataForSheet, { header: cleanedHeadersToUse });

        if (dataForSheet.length > 0 && ws['!ref']) {
            const range = XLSX.utils.decode_range(ws['!ref']);
            ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: range.e }) };
        }

        XLSX.utils.book_append_sheet(wb, ws, safeSheetName);
        XLSX.writeFile(wb, `PV_SubTeam_Report_${safeSheetName}.xlsx`);

    } else {
        // Logic for a single Claims Team (Full Data - All columns)
        const team = teamName;
        const claims = activeClaims.filter(c => c.assignedTeam === team);
        
        if (claims.length === 0) {
            alert(`No active claims found for Team: ${team}.`);
            return;
        }

        const headersToUse = globalHeaders;
        const cleanedHeadersToUse = cleanHeaders(headersToUse);

        const dataForSheet = claims.map(claim => {
            const row = {};
            headersToUse.forEach(originalHeader => {
                row[cleanHeaders([originalHeader])[0]] = claim[originalHeader];
            });
            return row;
        });

        const ws = XLSX.utils.json_to_sheet(dataForSheet, { header: cleanedHeadersToUse });

        // Apply formatting and AutoFilter
        if (dataForSheet.length > 0 && ws['!ref']) {
            const range = XLSX.utils.decode_range(ws['!ref']);
            ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: range.e }) };

            cleanedHeadersToUse.forEach((header, colIndex) => {
                const originalHeader = headersToUse.find(h => cleanHeaders([h])[0] === header);
                if (DATE_COLUMNS.includes(originalHeader)) {
                    for (let i = range.s.r + 1; i <= range.e.r; ++i) { 
                        const cellRef = XLSX.utils.encode_cell({c: colIndex, r: i});
                        const cell = ws[cellRef];
                        if (cell && cell.v) {
                            const dateCell = formatAsDateCell(cell.v);
                            if (dateCell.t === 'n') {
                                cell.v = dateCell.v;
                                cell.t = dateCell.t;
                                cell.z = dateCell.z;
                            }
                        }
                    }
                }
            });
        }

        const safeSheetName = team.replace(/\s\(.*\)/m, '').replace(/[^a-zA-Z0-9]/g, '_').substring(0, 30).trim();
        XLSX.utils.book_append_sheet(wb, ws, "All Claims Data");
        XLSX.writeFile(wb, `${safeSheetName}_All_Data_Report.xlsx`);
    }
}
// --- END downloadReport IMPLEMENTATION ---


function countClaimsByStatusAndDSNP(claims) {
    const counts = {
        overall: { active: {}, prebatch: {}, combined: {} },
        dsnp: { active: {}, prebatch: {}, combined: {} },
        nondsnp: { active: {}, prebatch: {}, combined: {} }
    };
    const agingOrderMap = { '0-20 Queue': '0-20', '21-27 Priority': '21-27', '28-30 Critical': '28-30', '31+ Backlog': '31+' };

    claims.forEach(claim => {
        // Only count if assigned to a working team (exclude UNASSIGNED from this analysis)
        if(claim.assignedTeam === CLAIMS_FUNCTIONS.UNASSIGNED) return;

        const ageBucket = agingOrderMap[claim.ageBucket] || claim.ageBucket;
        const networkStatus = claim.NetworkStatus.toUpperCase().trim();
        const isPar = networkStatus === 'IN NETWORK'; 
        const statusKey = claim.isPrebatch ? 'prebatch' : 'active';
        const dsnpKey = claim['DSNP or Non DSNP'].toUpperCase() === 'DSNP' ? 'dsnp' : 'nondsnp';
        
        const increment = (target, aging, par) => {
            target.total = (target.total || 0) + 1;
            target[aging] = target[aging] || { Par: 0, NonPar: 0, Total: 0 };
            target[aging].Total++;
            if (par) {
                target.ParTotal = (target.ParTotal || 0) + 1;
                target[aging].Par++;
            } else {
                target.NonParTotal = (target.NonParTotal || 0) + 1;
                target[aging].NonPar++;
            }
        };

        increment(counts.overall[statusKey], ageBucket, isPar);
        increment(counts.overall.combined, ageBucket, isPar);
        increment(counts[dsnpKey][statusKey], ageBucket, isPar);
        increment(counts[dsnpKey].combined, ageBucket, isPar);
    });
    
    return counts;
}

function generateOverallFocusHtml(counts) {
    const agingOrder = ['0-20', '21-27', '28-30', '31+'];
    const sectionOrder = ['overall', 'dsnp', 'nondsnp'];
    const statusOrder = ['combined', 'active', 'prebatch'];
    
    let html = '<h2 class="overall-focus-header">Overall Focus: Claim Count Analysis (Excluding Needs Assignment)</h2>';

    for (let statusKey of statusOrder) {
        let statusTitle;
        if (statusKey === 'combined') {
            statusTitle = 'Combined Claim Counts (Active + Prebatch)';
        } else if (statusKey === 'active') {
            statusTitle = 'Active-Only Claim Counts Analysis';
        } else {
            statusTitle = 'Prebatch-Only Claim Counts Analysis';
        }
        
        html += `<h3 style="margin-top: ${statusKey === 'combined' ? '15px' : '30px'};">${statusTitle}</h3>`; 
        
        let tablesHtml = '';
        
        for (let sectionKey of sectionOrder) {
            const data = counts[sectionKey][statusKey];
            let sectionTitle;
            let filterTitle;
            if (sectionKey === 'overall') {
                sectionTitle = statusKey === 'combined' ? 'Overall Combined Counts' : (statusKey === 'active' ? 'Active Counts (Overall)' : 'Prebatch Counts (Overall)');
                filterTitle = statusKey === 'combined' ? 'Filters: All Clean Claims (Overall Only)' : (statusKey === 'active' ? 'Filters: Clean Claims, Overall Only, All but Prebatch' : 'Filters: Clean Claims, Overall Only, Prebatch Only');
            } else if (sectionKey === 'dsnp') {
                sectionTitle = statusKey === 'combined' ? 'Combined Counts (DSNP)' : (statusKey === 'active' ? 'Active Counts (DSNP)' : 'Prebatch Counts (DSNP)');
                filterTitle = statusKey === 'combined' ? 'Filters: All Clean Claims (DSNP Only)' : (statusKey === 'active' ? 'Filters: Clean Claims, DSNP Only, All but Prebatch' : 'Filters: Clean Claims, DSNP Only, Prebatch Only');
            } else {
                sectionTitle = statusKey === 'combined' ? 'Combined Counts (Non DSNP)' : (statusKey === 'active' ? 'Active Counts (Non DSNP)' : 'Prebatch Counts (Non DSNP)');
                filterTitle = statusKey === 'combined' ? 'Filters: All Clean Claims (Non DSNP Only)' : (statusKey === 'active' ? 'Filters: Clean Claims, Non DSNP Only, All but Prebatch' : 'Filters: Clean Claims, Non DSNP Only, Prebatch Only');
            }
            
            if (data.total === undefined || data.total === 0) { continue; }

            const totalPar = data.ParTotal || 0;
            const totalNonPar = data.NonParTotal || 0;
            const grandTotal = data.total;

            const p31 = data['31+'] || { Par: 0, NonPar: 0 };
            const par31Percent = totalPar > 0 ? ((p31.Par / totalPar) * 100).toFixed(1) + '%' : '0.0%';
            const nonPar31Percent = totalNonPar > 0 ? ((p31.NonPar / totalNonPar) * 100).toFixed(1) + '%' : '0.0%';
            const total31Percent = grandTotal > 0 ? (((p31.Par + p31.NonPar) / grandTotal) * 100).toFixed(1) + '%' : '0.0%';

            let tableGroupHtml = `
                <div class="pdf-page-break-avoid">
                    <h4>${sectionTitle}</h4>
                    <p style="font-size: 0.9em; margin-bottom: 5px;">${filterTitle}</p>
                    <table class="overall-focus-table">
                        <thead>
                            <tr class="overall-focus-subheader">
                                <th style="width: 15%;">Aging</th>
                                <th style="width: 25%; text-align: right;">Par</th>
                                <th style="width: 25%; text-align: right;">Non Par</th>
                                <th style="width: 35%; text-align: right;">Grand Total</th>
                            </tr>
                        </thead>
                        <tbody>
            `;

            for (const age of agingOrder) {
                const rowData = data[age] || { Par: 0, NonPar: 0, Total: 0 };
                tableGroupHtml += `
                    <tr>
                        <td>${age}</td>
                        <td align="right">${rowData.Par.toLocaleString()}</td>
                        <td align="right">${rowData.NonPar.toLocaleString()}</td>
                        <td align="right">${rowData.Total.toLocaleString()}</td>
                    </tr>
                `;
            }

            tableGroupHtml += `
                <tr>
                    <td>31+ Percentage</td>
                    <td align="right" class="overall-focus-highlight">${par31Percent}</td>
                    <td align="right" class="overall-focus-highlight">${nonPar31Percent}</td>
                    <td align="right" class="overall-focus-highlight">${total31Percent}</td>
                </tr>
            `;

            tableGroupHtml += `
                <tr class="overall-focus-highlight">
                    <td>Grand Total</td>
                    <td align="right">${totalPar.toLocaleString()}</td>
                    <td align="right">${totalNonPar.toLocaleString()}</td>
                    <td align="right">${grandTotal.toLocaleString()}</td>
                </tr>
            `;
            
            tableGroupHtml += '</tbody></table></div>';
            tablesHtml += tableGroupHtml;
        }
        
        html += tablesHtml; 
    }
    
    return html;
}

function updateReassignmentStatusDisplay() {
    const rulesCount = Object.keys(globalReassignmentRules).length;
    const statusElement = document.getElementById('reassignment-status');
    if (!statusElement) return;

    if (rulesCount > 0) {
        statusElement.innerText = `Custom rules loaded automatically: ${rulesCount} rule(s) waiting for file analysis.`;
        statusElement.style.color = 'orange'; 
        document.getElementById('reassignment-section').style.display = 'block'; // Ensure section is visible if rules exist
    } else {
        statusElement.innerText = 'No custom rules currently applied.';
        statusElement.style.color = 'black';
    }
}

function displayResults(groupingMap, totalActiveClaimsCount) {
    const resultsDiv = document.getElementById('results');
    const sortedGroups = Array.from(groupingMap.values()).sort((a, b) => b.count - a.count);

    const teamTotals = {};
    const pvSubTeamTotals = {};
    let claimsTeamTotalLocal = 0;
    let unassignedTotal = 0;
    
    sortedGroups.forEach(group => {
        teamTotals[group.team] = (teamTotals[group.team] || 0) + group.count;
        if (group.team === CLAIMS_FUNCTIONS.CLAIMS) {
            claimsTeamTotalLocal += group.count;
        } else if (group.team === CLAIMS_FUNCTIONS.PV) {
            if (group.subTeam) {
                pvSubTeamTotals[group.subTeam] = (pvSubTeamTotals[group.subTeam] || 0) + group.count;
            }
        } else if (group.team === CLAIMS_FUNCTIONS.UNASSIGNED) {
             unassignedTotal += group.count;
        }
    });
    
    pvTotalGlobal = teamTotals[CLAIMS_FUNCTIONS.PV] || 0;
    pvSubTeamTotalsGlobal = pvSubTeamTotals;
    // Manual update of PV Sub-Team Names
    pvSubTeamNamesGlobal = [
        'Provider/Vendor Creation',
        'W9/Validation/COB',
        'Contract/Network Issues',
        'Pay-to Provider Issues',
        'Pricing/PBP/Other',
        'PV Team (Uncategorized)'
    ].filter(name => pvSubTeamTotals[name] > 0);
    // END Manual update
    
    claimsTeamTotalLocal = teamTotals[CLAIMS_FUNCTIONS.CLAIMS] || 0;
    claimsTeamTotalGlobal = claimsTeamTotalLocal; 
    
    
    // --- OVERALL FOCUS TABLES (TOP SECTION) ---
    const allCounts = countClaimsByStatusAndDSNP(globalAllClaims);
    const overallFocusHtml = generateOverallFocusHtml(allCounts);
    let contentHtml = overallFocusHtml;

    // Add Download PDF Button
    contentHtml = `<button class="pdf-main-btn no-pdf" onclick="downloadPDF()">Download Full Report PDF</button>` + contentHtml;


    // --- TOP-LEVEL SUMMARY TABLE (SECOND SECTION) ---
    let summaryHtml = `
        <h2 style="color: #007bff;">Top-Level Workload Summary (${totalActiveClaimsCount.toLocaleString()} Active Claims)</h2>
        <div class="pdf-page-break-avoid">
        <table>
            <thead>
                <tr class="SUMMARY">
                    <th style="width: 30%;">Team</th>
                    <th style="width: 15%; text-align: center;">Total Active Claims</th>
                    <th style="width: 15%;">Percentage</th>
                    <th style="width: 40%;" class="no-pdf">Download Report</th>
                </tr>
            </thead>
            <tbody>
    `;
    
    const topLevelTotals = {
        [CLAIMS_FUNCTIONS.PV]: pvTotalGlobal,
        [CLAIMS_FUNCTIONS.CLAIMS]: claimsTeamTotalLocal,
        [CLAIMS_FUNCTIONS.UNASSIGNED]: unassignedTotal
    };

    // Determine if rules have been applied to update the status message
    const rulesMessage = Object.keys(globalReassignmentRules).length > 0 ? 
                         `Custom rules applied (${Object.keys(globalReassignmentRules).length} rules)` : 
                         'No custom rules currently applied.';
    document.getElementById('reassignment-status').innerText = rulesMessage;
    document.getElementById('reassignment-status').style.color = Object.keys(globalReassignmentRules).length > 0 ? 'green' : 'black';


    // NEW: Include UNASSIGNED in the summary table
    let teamOrder = [CLAIMS_FUNCTIONS.PV, CLAIMS_FUNCTIONS.CLAIMS, CLAIMS_FUNCTIONS.UNASSIGNED];
    teamOrder.forEach(team => {
        const count = topLevelTotals[team] || 0;
        const percentage = totalActiveClaimsCount > 0 ? ((count / totalActiveClaimsCount) * 100).toFixed(1) : "0.0";
        let teamClass = 'UNASSIGNED';
        let downloadButton = `<p style="font-size: 0.8em; color: #dc3545; font-weight: bold;">Assignment Required</p>`;

        if (team === CLAIMS_FUNCTIONS.PV) {
            teamClass = 'PV';
            downloadButton = `
                <button class="pv-workbook-btn no-pdf" onclick="downloadPvXlsxWorkbook()">PV Workload (No PHI)</button>
                <button class="pv-full-data-btn no-pdf" onclick="downloadPvFullDataWorkbook()" style="margin-top: 5px; font-size: 14px; padding: 8px 10px;">PV Full Data (with PHI)</button>
            `;
        } else if (team === CLAIMS_FUNCTIONS.CLAIMS) {
            teamClass = 'CLAIMS';
            downloadButton = `<button class="claims-workbook-btn no-pdf" onclick="downloadClaimsXlsxWorkbook()">Claims Master Workbook</button>`;
        }


        summaryHtml += `
            <tr class="${teamClass}">
                <td style="font-weight: bold;">${team.replace(/\s\(.*\)/m, '')}</td>
                <td class="count">${count.toLocaleString()}</td>
                <td>${percentage}%</td>
                <td class="no-pdf">${downloadButton}</td>
            </tr>
        `;
    });
    summaryHtml += '</tbody></table></div>';


    // --- PV INTERNAL BREAKDOWN ---
    summaryHtml += `
        <h3 style="margin-top: 20px;">PV Team (Provider Ops) Internal Reports (Your Tabs):</h3>
        <div class="pdf-page-break-avoid">
        <table>
            <thead>
                <tr class="SUMMARY">
                    <th style="width: 30%;">Report / Tab Name</th>
                    <th style="width: 15%; text-align: center;">Total Active Claims</th>
                    <th style="width: 15%;">Percentage of PV Work</th>
                    <th style="width: 40%;" class="no-pdf">Download (Included in Workbook)</th>
                </tr>
            </thead>
            <tbody>
    `;
    
    pvSubTeamNamesGlobal.forEach(subTeam => {
        const count = pvSubTeamTotalsGlobal[subTeam] || 0;
        const percentage = pvTotalGlobal > 0 ? ((count / pvTotalGlobal) * 100).toFixed(1) : 0.0;
        
        summaryHtml += `
            <tr class="PV_SUB">
                <td style="font-weight: bold;">${subTeam}</td>
                <td class="count">${count.toLocaleString()}</td>
                <td>${percentage}%</td>
                <td class="no-pdf"><button class="download-btn no-pdf" onclick="downloadReport('${subTeam}', true)">Single XLSX</button></td>
            </tr>
        `;
    });
    summaryHtml += '</tbody></table></div>';


    // --- CLAIMS TEAM INTERNAL BREAKDOWN (PDF FIX: Wrapped everything in one container) ---
    // Define ageHeaders here to fix the "not defined" error
    let ageHeaders = ['0-20 Queue', '21-27 Priority', '28-30 Critical', '31+ Backlog'];

    summaryHtml += `
        <div class="pdf-page-break-avoid"> <!-- Wrap the entire claims breakdown section -->
            <h3 style="margin-top: 20px;">Claims Team Internal Breakdown:</h3>
            
            <h4 class="AgeType">Clean Age Aging (Since Ready to Pay - Active Claims Only)</h4>
            <table class="claims-aging-table">
                <thead>
                    <tr class="claims-summary-header">
                        <th style="width: 15%;">Internal Function</th>
                        <th style="width: 10%;">Total Claims</th>
                        ${ageHeaders.map(h => `<th style="width: 15%; text-align: center;">${h}</th>`).join('')}
                        <th style="width: 10%;" class="no-pdf">Download</th>
                    </tr>
                </thead>
                <tbody>
    `;
    
    // Only one team to display here now: Claims Team
    const team = CLAIMS_FUNCTIONS.CLAIMS;
    const count = teamTotals[team] || 0;
    const activeClaims = globalAllClaims.filter(c => !c.isPrebatch); 
    const claimsInFunc = activeClaims.filter(c => c.assignedTeam === team);
    const teamClass = 'CLAIMS';
    const teamNameShort = team.replace(/\s\(.*\)/m, '');

    const cleanAgeCounts = claimsInFunc.reduce((acc, claim) => {
        acc[claim.ageBucket] = (acc[claim.ageBucket] || 0) + 1;
        return acc;
    }, {});

    summaryHtml += `
        <tr class="${teamClass}">
            <td style="font-weight: bold;">${teamNameShort}</td>
            <td class="count">${count.toLocaleString()}</td>
            ${ageHeaders.map(bucket => {
                const ageCount = cleanAgeCounts[bucket] || 0;
                const ageClass = bucket.split(' ')[1]; 
                return `<td class="count ${ageClass}">${ageCount.toLocaleString()}</td>`;
            }).join('')}
            <td class="no-pdf"><button class="download-btn no-pdf" onclick="downloadReport('${team}')">XLSX</button></td>
        </tr>
    `;
    summaryHtml += '</tbody></table>';


    // --- CLAIMS TEAM INTERNAL BREAKDOWN - RECEIVED AGE TABLE ---
    summaryHtml += `
        <h4 class="AgeType" style="margin-top: 20px;">Received Age Aging (Since Received Date - Active Claims Only)</h4>
        <table class="claims-aging-table">
            <thead>
                <tr class="claims-summary-header">
                    <th style="width: 15%;">Internal Function</th>
                    <th style="width: 10%;">Total Claims</th>
                    ${ageHeaders.map(h => `<th style="width: 15%; text-align: center;">${h}</th>`).join('')}
                    <th style="width: 10%;" class="no-pdf">Download</th>
                </tr>
            </thead>
            <tbody>
    `;
    
    // Only one team to display here now: Claims Team
    const receivedAgeCounts = claimsInFunc.reduce((acc, claim) => { 
        acc[claim.receivedAgeBucket] = (acc[claim.receivedAgeBucket] || 0) + 1;
        return acc;
    }, {});
    
    summaryHtml += `
        <tr class="${teamClass}">
            <td style="font-weight: bold;">${teamNameShort}</td>
            <td class="count">${count.toLocaleString()}</td>
            ${ageHeaders.map(bucket => {
                const ageCount = receivedAgeCounts[bucket] || 0;
                const ageClass = bucket.split(' ')[1]; 
                return `<td class="count ${ageClass}">${ageCount.toLocaleString()}</td>`;
            }).join('')}
            <td class="no-pdf"><button class="download-btn no-pdf" onclick="downloadReport('${team}')">XLSX</button></td>
        </tr>
    `;
    summaryHtml += '</tbody></table>';
    summaryHtml += '</div>'; // Close the pdf-page-break-avoid div
    
    // Final Output
    resultsDiv.innerHTML = contentHtml + summaryHtml;
}

window.onload = function() {
    document.getElementById('analyzeButton').onclick = processData;
    document.getElementById('openReassignEditor').onclick = openReassignmentEditor;
    document.getElementById('applyReassignments').onclick = applyAssignmentsAndRedraw;

    // **NEW: Initial status update on load**
    updateReassignmentStatusDisplay(); 

    // Make the download and import/export functions globally accessible
    window.downloadClaimsXlsxWorkbook = downloadClaimsXlsxWorkbook;
    window.downloadPvFullDataWorkbook = downloadPvFullDataWorkbook; 
    window.downloadPvXlsxWorkbook = downloadPvXlsxWorkbook;
    window.downloadReport = downloadReport; // The newly implemented function
    window.downloadPDF = downloadPDF;
    window.exportAssignmentRules = exportAssignmentRules; 
    window.importAssignmentRules = importAssignmentRules; 
    window.clearAssignmentRules = clearAssignmentRules; // NEW function made globally accessible
};
