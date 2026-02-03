/**
 * Surplus Lines Tax API - Excel Custom Functions
 *
 * Custom functions for calculating surplus lines taxes directly in Excel.
 * All functions are available under the SLTAX namespace.
 *
 * These functions match exactly with the Google Sheets integration.
 *
 * Examples:
 *   =SLTAX.CALCULATE("Texas", 10000)
 *   =SLTAX.RATE("California")
 *   =SLTAX.DETAILS("New York", 25000)
 *
 * Get your API key: https://app.surpluslinesapi.com
 * Documentation: https://surpluslinesapi.com/excel/
 *
 * © Underwriters Technologies - https://undtec.com
 */

// ============================================================================
// CONFIGURATION
// ============================================================================

const API_BASE_URL = 'https://api.surpluslinesapi.com/v1';

/**
 * Get API key from Office settings or storage
 * @returns {string|null} The API key or null if not configured
 */
function getApiKey() {
    // Try to get from Office roaming settings
    if (typeof Office !== 'undefined' && Office.context && Office.context.roamingSettings) {
        const key = Office.context.roamingSettings.get('sltax_api_key');
        if (key) return key;
    }

    // Try localStorage as fallback
    if (typeof localStorage !== 'undefined') {
        const key = localStorage.getItem('sltax_api_key');
        if (key) return key;
    }

    return null;
}

/**
 * Set API key in Office settings
 * @param {string} apiKey - The API key to store
 */
function setApiKey(apiKey) {
    if (typeof Office !== 'undefined' && Office.context && Office.context.roamingSettings) {
        Office.context.roamingSettings.set('sltax_api_key', apiKey);
        Office.context.roamingSettings.saveAsync();
    }

    if (typeof localStorage !== 'undefined') {
        localStorage.setItem('sltax_api_key', apiKey);
    }
}

// ============================================================================
// CUSTOM FUNCTIONS
// ============================================================================

/**
 * Calculate surplus lines tax for a given state and premium amount.
 * @customfunction CALCULATE
 * @param {string} state State name (e.g., "Texas", "California") or abbreviation (e.g., "TX", "CA")
 * @param {number} premium Premium amount in dollars
 * @returns {Promise<number|string>} Total tax amount or error message
 * @example =SLTAX.CALCULATE("Texas", 10000)
 * Returns: 503
 */
async function calculate(state, premium) {
    // Input validation
    if (!state || state === '') {
        return 'ERROR: State is required';
    }

    if (!premium || premium <= 0) {
        return 'ERROR: Premium must be greater than 0';
    }

    const apiKey = getApiKey();
    if (!apiKey) {
        return 'ERROR: Please configure your API key in Settings';
    }

    try {
        const response = await fetch(`${API_BASE_URL}/calculate`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'X-API-Key': apiKey
            },
            body: JSON.stringify({
                state: state,
                premium: parseFloat(premium)
            })
        });

        const data = await response.json();

        if (!response.ok || data.error === true || !data.success) {
            return `ERROR: ${data.message || data.error || 'API error'}`;
        }

        return data.total_tax || 0;

    } catch (error) {
        return `ERROR: ${error.message}`;
    }
}

/**
 * Get detailed tax breakdown including state, premium, total tax, and total due.
 * @customfunction DETAILS
 * @param {string} state State name or abbreviation
 * @param {number} premium Premium amount in dollars
 * @param {boolean} [multiline] If true, returns data in multiple rows; if false (default), returns in one row
 * @returns {Promise<any[][]>} Array with complete tax breakdown
 * @example =SLTAX.DETAILS("California", 10000)
 * Returns (one row): California | 10000 | 318 | 10318
 * @example =SLTAX.DETAILS("California", 10000, TRUE)
 * Returns (multiple rows): State: California, Premium: 10000, Total Tax: 318, Total Due: 10318
 */
async function details(state, premium, multiline) {
    multiline = multiline || false;

    // Input validation
    if (!state || state === '') {
        return [['ERROR: State is required']];
    }

    if (!premium || premium <= 0) {
        return [['ERROR: Premium must be greater than 0']];
    }

    const apiKey = getApiKey();
    if (!apiKey) {
        return [['ERROR: Configure API key']];
    }

    try {
        const response = await fetch(`${API_BASE_URL}/calculate`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'X-API-Key': apiKey
            },
            body: JSON.stringify({
                state: state,
                premium: parseFloat(premium)
            })
        });

        const data = await response.json();

        if (!response.ok || data.error === true || !data.success) {
            return [[`ERROR: ${data.message || data.error || 'API error'}`]];
        }

        if (multiline) {
            // Return as vertical array (multiple rows, 1 column)
            return [
                [`State: ${data.state}`],
                [`Premium: ${data.premium}`],
                [`Total Tax: ${data.total_tax}`],
                [`Total Due: ${data.total_due}`]
            ];
        } else {
            // Return as horizontal array (1 row, multiple columns)
            return [[
                data.state || state,
                data.premium || premium,
                data.total_tax || 0,
                data.total_due || 0
            ]];
        }

    } catch (error) {
        return [[`ERROR: ${error.message}`]];
    }
}

/**
 * Calculate tax and return premium, total_tax, and total_due in a row.
 * @customfunction WITHPREMIUM
 * @param {string} state State name or abbreviation
 * @param {number} premium Premium amount in dollars
 * @returns {Promise<any[][]>} Array with [premium, total_tax, total_due]
 * @example =SLTAX.WITHPREMIUM("California", 10000)
 * Returns: 10000 | 318 | 10318
 */
async function withPremium(state, premium) {
    // Input validation
    if (!state || state === '') {
        return [['ERROR: State is required', '', '']];
    }

    if (!premium || premium <= 0) {
        return [['ERROR: Premium must be greater than 0', '', '']];
    }

    const apiKey = getApiKey();
    if (!apiKey) {
        return [['ERROR: Configure API key', '', '']];
    }

    try {
        const response = await fetch(`${API_BASE_URL}/calculate`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'X-API-Key': apiKey
            },
            body: JSON.stringify({
                state: state,
                premium: parseFloat(premium)
            })
        });

        const data = await response.json();

        if (!response.ok || data.error === true || !data.success) {
            return [[`ERROR: ${data.message || data.error || 'API error'}`, '', '']];
        }

        // Return as horizontal array (1 row, 3 columns)
        return [[
            parseFloat(premium),
            data.total_tax || 0,
            data.total_due || 0
        ]];

    } catch (error) {
        return [[`ERROR: ${error.message}`, '', '']];
    }
}

/**
 * Get the surplus lines tax rate for a specific state.
 * @customfunction RATE
 * @param {string} state State name (e.g., "Texas", "California") or abbreviation (e.g., "TX", "CA")
 * @returns {Promise<number|string>} Tax rate as a percentage (e.g., 4.85 for 4.85%)
 * @example =SLTAX.RATE("California")
 * Returns: 3
 */
async function rate(state) {
    // Input validation
    if (!state || state === '') {
        return 'ERROR: State is required';
    }

    const apiKey = getApiKey();
    if (!apiKey) {
        return 'ERROR: Please configure your API key in Settings';
    }

    try {
        const response = await fetch(`${API_BASE_URL}/rates?state=${encodeURIComponent(state)}`, {
            method: 'GET',
            headers: {
                'X-API-Key': apiKey
            }
        });

        const data = await response.json();

        if (!response.ok || data.error === true || !data.success) {
            return `ERROR: ${data.message || data.error || 'API error'}`;
        }

        // Response format: {"success":true,"data":[{"state":"California","tax_rate":"3%",...}]}
        const rates = data.data || [];
        if (!Array.isArray(rates) || rates.length === 0) {
            return `ERROR: State '${state}' not found`;
        }

        const taxRateStr = rates[0].tax_rate || '0%';

        // Parse rate string (e.g., "3%" -> 3)
        if (typeof taxRateStr === 'string') {
            const match = taxRateStr.match(/[\d.]+/);
            if (match) {
                return parseFloat(match[0]);
            }
        } else if (typeof taxRateStr === 'number') {
            return taxRateStr;
        }

        return 0;

    } catch (error) {
        return `ERROR: ${error.message}`;
    }
}

/**
 * Get list of all supported states.
 * @customfunction STATES
 * @returns {Promise<string[][]>} Array of state names (vertical)
 * @example =SLTAX.STATES()
 * Returns: Alabama, Alaska, Arizona, ...
 */
async function states() {
    const apiKey = getApiKey();
    if (!apiKey) {
        return [['ERROR: Please configure your API key in Settings']];
    }

    try {
        const response = await fetch(`${API_BASE_URL}/states`, {
            method: 'GET',
            headers: {
                'X-API-Key': apiKey
            }
        });

        const data = await response.json();

        if (!response.ok || data.error === true || !data.success) {
            return [[`ERROR: ${data.message || data.error || 'API error'}`]];
        }

        // Handle different response structures
        const stateList = data.states || data.data || [];
        if (!Array.isArray(stateList)) {
            return [['ERROR: Invalid response format']];
        }

        // Return as vertical array (one state per row)
        return stateList.map(s => [typeof s === 'string' ? s : (s.name || s)]);

    } catch (error) {
        return [[`ERROR: ${error.message}`]];
    }
}

/**
 * Get tax rates for all states.
 * Returns a 2-column array: state name and tax rate.
 * @customfunction RATES
 * @returns {Promise<any[][]>} Array with [state, rate] for each row
 * @example =SLTAX.RATES()
 * Returns: Texas | 4.85, California | 3, ...
 */
async function rates() {
    const apiKey = getApiKey();
    if (!apiKey) {
        return [['ERROR: Please configure your API key in Settings', '']];
    }

    try {
        const response = await fetch(`${API_BASE_URL}/rates`, {
            method: 'GET',
            headers: {
                'X-API-Key': apiKey
            }
        });

        const data = await response.json();

        if (!response.ok || data.error === true || !data.success) {
            return [[`ERROR: ${data.message || data.error || 'API error'}`, '']];
        }

        // Response format: {"success":true,"data":[{"state":"Alabama","tax_rate":"6%",...}]}
        const ratesList = data.data || [];
        if (!Array.isArray(ratesList)) {
            return [['ERROR: Invalid response format', '']];
        }

        // Return as 2-column array: [state, rate]
        return ratesList.map(item => {
            const stateName = item.state || 'Unknown';
            const taxRateStr = item.tax_rate || '0%';

            // Parse rate string (e.g., "3%" -> 3)
            let ratePercent = 0;
            if (typeof taxRateStr === 'string') {
                const match = taxRateStr.match(/[\d.]+/);
                if (match) {
                    ratePercent = parseFloat(match[0]);
                }
            } else if (typeof taxRateStr === 'number') {
                ratePercent = taxRateStr;
            }

            return [stateName, ratePercent];
        });

    } catch (error) {
        return [[`ERROR: ${error.message}`, '']];
    }
}

/**
 * Get historical tax rate for a specific state and date.
 * @customfunction HISTORICALRATE
 * @param {string} state State name (e.g., "Texas", "Iowa")
 * @param {any} [date] Date in YYYY-MM-DD format or Excel date serial number
 * @returns {Promise<number|string>} Tax rate as a percentage
 * @example =SLTAX.HISTORICALRATE("Iowa", "2025-06-15")
 * Returns: 0.95
 */
async function historicalRate(state, date) {
    // Input validation
    if (!state || state === '') {
        return 'ERROR: State is required';
    }

    if (!date || date === '') {
        return 'ERROR: Date is required (YYYY-MM-DD format)';
    }

    const apiKey = getApiKey();
    if (!apiKey) {
        return 'ERROR: Please configure your API key in Settings';
    }

    // Handle Excel date serial numbers
    let dateStr = date;
    if (typeof date === 'number') {
        // Convert Excel date serial to JS Date
        const excelEpoch = new Date(1899, 11, 30);
        const jsDate = new Date(excelEpoch.getTime() + date * 86400000);
        dateStr = jsDate.toISOString().split('T')[0];
    }

    // Validate date format
    const dateRegex = /^\d{4}-\d{2}-\d{2}$/;
    if (typeof dateStr === 'string' && !dateRegex.test(dateStr)) {
        return 'ERROR: Date must be in YYYY-MM-DD format';
    }

    try {
        const url = `${API_BASE_URL}/historical-rates?state=${encodeURIComponent(state)}&date=${encodeURIComponent(dateStr)}`;
        const response = await fetch(url, {
            method: 'GET',
            headers: {
                'X-API-Key': apiKey
            }
        });

        const data = await response.json();

        if (!response.ok || data.error === true || !data.success) {
            return `ERROR: ${data.message || data.error || 'API error'}`;
        }

        // Response format: {"success":true,"rate":{"tax_rate":0.0485,...},...}
        const rateData = data.rate || {};
        const taxRate = rateData.tax_rate;

        if (taxRate !== undefined && taxRate !== null) {
            // If it's a decimal (e.g., 0.0485), convert to percentage
            if (taxRate < 1) {
                return taxRate * 100;
            }
            return taxRate;
        }

        return 0;

    } catch (error) {
        return `ERROR: ${error.message}`;
    }
}

/**
 * Get detailed historical tax rate information for a specific state and date.
 * @customfunction HISTORICALRATEDETAILS
 * @param {string} state State name (e.g., "Texas", "Iowa")
 * @param {any} date Date in YYYY-MM-DD format or Excel date serial number
 * @param {boolean} [multiline] If true, returns data in multiple rows; if false (default), returns in one row
 * @returns {Promise<any[][]>} Array with complete historical rate breakdown
 * @example =SLTAX.HISTORICALRATEDETAILS("Texas", "2024-06-15")
 * Returns: Texas | 2024-06-15 | 4.85% | 0.18% | ...
 */
async function historicalRateDetails(state, date, multiline) {
    multiline = multiline || false;

    // Input validation
    if (!state || state === '') {
        return [['ERROR: State is required']];
    }

    if (!date || date === '') {
        return [['ERROR: Date is required (YYYY-MM-DD format)']];
    }

    const apiKey = getApiKey();
    if (!apiKey) {
        return [['ERROR: Configure API key']];
    }

    // Handle Excel date serial numbers
    let dateStr = date;
    if (typeof date === 'number') {
        const excelEpoch = new Date(1899, 11, 30);
        const jsDate = new Date(excelEpoch.getTime() + date * 86400000);
        dateStr = jsDate.toISOString().split('T')[0];
    }

    // Validate date format
    const dateRegex = /^\d{4}-\d{2}-\d{2}$/;
    if (typeof dateStr === 'string' && !dateRegex.test(dateStr)) {
        return [['ERROR: Date must be in YYYY-MM-DD format']];
    }

    try {
        const url = `${API_BASE_URL}/historical-rates?state=${encodeURIComponent(state)}&date=${encodeURIComponent(dateStr)}`;
        const response = await fetch(url, {
            method: 'GET',
            headers: {
                'X-API-Key': apiKey
            }
        });

        const data = await response.json();

        if (!response.ok || data.error === true || !data.success) {
            return [[`ERROR: ${data.message || data.error || 'API error'}`]];
        }

        const rate = data.rate || {};

        if (multiline) {
            // Return as vertical array (multiple rows, 1 column)
            return [
                [`State: ${state}`],
                [`Date: ${dateStr}`],
                [`Tax Rate: ${rate.tax_rate || 'N/A'}`],
                [`Stamping Fee: ${rate.stamping_fee || 'N/A'}`],
                [`Filing Fee: ${rate.filing_fee || 'N/A'}`],
                [`Service Fee: ${rate.service_fee || 'N/A'}`],
                [`Surcharge: ${rate.surcharge || 'N/A'}`],
                [`Regulatory Fee: ${rate.regulatory_fee || 'N/A'}`],
                [`Fire Marshal Tax: ${rate.fire_marshal_tax || 'N/A'}`],
                [`SLAS Clearinghouse Fee: ${rate.slas_clearinghouse_fee || 'N/A'}`],
                [`Flat Fee: ${rate.flat_fee || 'N/A'}`],
                [`Effective From: ${rate.effective_from || 'N/A'}`],
                [`Effective To: ${rate.effective_to || 'N/A'}`],
                [`Legislative Source: ${rate.legislative_source || 'N/A'}`],
                [`Confidence: ${rate.confidence || 'N/A'}`]
            ];
        } else {
            // Return as horizontal array (1 row, multiple columns)
            return [[
                state,
                dateStr,
                rate.tax_rate || null,
                rate.stamping_fee || null,
                rate.filing_fee || null,
                rate.service_fee || null,
                rate.surcharge || null,
                rate.regulatory_fee || null,
                rate.fire_marshal_tax || null,
                rate.slas_clearinghouse_fee || null,
                rate.flat_fee || null,
                rate.effective_from || null,
                rate.effective_to || null,
                rate.legislative_source || null,
                rate.confidence || null
            ]];
        }

    } catch (error) {
        return [[`ERROR: ${error.message}`]];
    }
}

/**
 * Get detailed tax rate information for all states.
 * @customfunction RATESDETAILS
 * @returns {Promise<any[][]>} Array with detailed rate information (one state per row)
 * @example =SLTAX.RATESDETAILS()
 * Returns: State | Tax Rate | Stamping Fee | Filing Fee | ...
 */
async function ratesDetails() {
    const apiKey = getApiKey();
    if (!apiKey) {
        return [['ERROR: Please configure your API key in Settings']];
    }

    try {
        const response = await fetch(`${API_BASE_URL}/rates`, {
            method: 'GET',
            headers: {
                'X-API-Key': apiKey
            }
        });

        const data = await response.json();

        if (!response.ok || data.error === true || !data.success) {
            return [[`ERROR: ${data.message || data.error || 'API error'}`]];
        }

        // Response format: {"success":true,"data":[{"state":"Alabama","tax_rate":"6%","stamping_fee":"0.25%",...}]}
        const ratesList = data.data || [];
        if (!Array.isArray(ratesList)) {
            return [['ERROR: Invalid response format']];
        }

        // Return each state as a row with all API fields in separate columns
        return ratesList.map(item => {
            return [
                item.state || 'Unknown',
                item.tax_rate || null,
                item.stamping_fee || null,
                item.filing_fee || null,
                item.service_fee || null,
                item.surcharge || null,
                item.regulatory_fee || null,
                item.fire_marshal_tax || null,
                item.slas_clearinghouse_fee || null,
                item.flat_fee || null,
                item.legislative_source || null
            ];
        });

    } catch (error) {
        return [[`ERROR: ${error.message}`]];
    }
}

// ============================================================================
// FUNCTION REGISTRATION
// ============================================================================

// Register functions with Office if available
if (typeof CustomFunctions !== 'undefined') {
    CustomFunctions.associate('CALCULATE', calculate);
    CustomFunctions.associate('DETAILS', details);
    CustomFunctions.associate('WITHPREMIUM', withPremium);
    CustomFunctions.associate('RATE', rate);
    CustomFunctions.associate('STATES', states);
    CustomFunctions.associate('RATES', rates);
    CustomFunctions.associate('HISTORICALRATE', historicalRate);
    CustomFunctions.associate('HISTORICALRATEDETAILS', historicalRateDetails);
    CustomFunctions.associate('RATESDETAILS', ratesDetails);
}

// Export for testing
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        calculate,
        details,
        withPremium,
        rate,
        states,
        rates,
        historicalRate,
        historicalRateDetails,
        ratesDetails,
        getApiKey,
        setApiKey
    };
}
