/**
 * Surplus Lines Tax API - Excel Custom Functions v2.0
 *
 * Simplified custom function for calculating surplus lines taxes and getting rates
 * directly in Excel.
 *
 * Get your API key: https://app.surpluslinesapi.com
 * Documentation: https://surpluslinesapi.com/excel/
 *
 * © Underwriters Technologies - https://undtec.com
 */

// ============================================================================
// CONFIGURATION
// ============================================================================

const API_BASE_URL = 'https://n8n.undtec.com/webhook/slapi/v1';

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
// MAIN FUNCTION
// ============================================================================

/**
 * Calculate surplus lines tax or get tax rates for a state (v2.0).
 *
 * This function supports two calculation types:
 * - "Tax": Calculate tax amount for a given premium
 * - "Rate": Get detailed tax rates for a state
 *
 * @customfunction SLAPI
 * @param {string} calculationType - "Tax" or "Rate"
 * @param {string} effectiveDate - Date in YYYY-MM-DD format (optional, use "" for current rates)
 * @param {string} stateCode - State name (e.g., "Florida") or code (e.g., "FL")
 * @param {number} [premiumAmount] - Premium amount in dollars (required for "Tax", ignored for "Rate")
 * @returns {Promise<any[][]>} 2D array with results (vertical format: field name | value)
 *
 * @example =SLTAX.SLAPI("Tax", "", "Florida", 10000)
 * Returns:
 *   Base Tax      | 494
 *   Stamping Fee  | 6
 *
 * @example =SLTAX.SLAPI("Tax", "2020-01-01", "Texas", 10000)
 * Returns (if historical data not found):
 *   Base Tax      | 485
 *   Stamping Fee  | 5
 *   ⚠️ Notice     | No historical data available for 2020-01-01
 *   Rates From    | current
 *
 * @example =SLTAX.SLAPI("Rate", "", "Florida")
 * Returns:
 *   tax_rate                  | 4.94%
 *   stamping_fee             |
 *   filing_fee               |
 *   service_fee              | 0.06%
 *   surcharge                |
 *   regulatory_fee           |
 *   fire_marshal_tax         |
 *   slas_clearinghouse_fee   |
 *   flat_fee                 |
 *
 * @example =SLTAX.SLAPI("Rate", "2020-01-01", "Texas")
 * Returns (if historical data not found):
 *   tax_rate                  | 4.85%
 *   stamping_fee             | 0.05%
 *   filing_fee               |
 *   service_fee              |
 *   surcharge                |
 *   regulatory_fee           |
 *   fire_marshal_tax         |
 *   slas_clearinghouse_fee   |
 *   flat_fee                 |
 *   ⚠️ Notice                 | No historical data available for 2020-01-01
 *   Rates From               | current
 */
async function slapi(calculationType, effectiveDate, stateCode, premiumAmount) {
    // Input validation
    if (!calculationType || calculationType === '') {
        return [['ERROR: Calculation Type is required (use "Tax" or "Rate")']];
    }

    const calcType = calculationType.toString().trim().toLowerCase();
    if (calcType !== 'tax' && calcType !== 'rate') {
        return [['ERROR: Calculation Type must be "Tax" or "Rate"']];
    }

    const apiKey = getApiKey();
    if (!apiKey) {
        return [['ERROR: Please configure your API key in Settings']];
    }

    if (!stateCode || stateCode === '') {
        return [['ERROR: State Code is required']];
    }

    // Handle effectiveDate - empty string means current rates
    const dateParam = (effectiveDate && effectiveDate !== '') ? effectiveDate.toString().trim() : '';

    // Validate date format if provided
    if (dateParam !== '') {
        // Convert Excel date serial to string if needed
        let dateStr = dateParam;
        if (typeof dateParam === 'number') {
            const excelEpoch = new Date(1899, 11, 30);
            const jsDate = new Date(excelEpoch.getTime() + dateParam * 86400000);
            dateStr = jsDate.toISOString().split('T')[0];
        }

        const dateRegex = /^\d{4}-\d{2}-\d{2}$/;
        if (!dateRegex.test(dateStr)) {
            return [['ERROR: Date must be in YYYY-MM-DD format or empty string']];
        }
    }

    try {
        if (calcType === 'tax') {
            return await calculateTax(apiKey, stateCode, premiumAmount, dateParam);
        } else {
            return await getRate(apiKey, stateCode, dateParam);
        }
    } catch (error) {
        return [[`ERROR: ${error.message}`]];
    }
}

// ============================================================================
// HELPER FUNCTIONS (Internal - Do not call directly)
// ============================================================================

/**
 * Calculate tax and return Base Tax and Stamping Fee
 * @private
 */
async function calculateTax(apiKey, stateCode, premiumAmount, effectiveDate) {
    // Validate premium
    if (!premiumAmount || premiumAmount <= 0) {
        return [['ERROR: Premium must be greater than 0']];
    }

    // Build request payload
    const payload = {
        state: stateCode.toString().trim(),
        premium: parseFloat(premiumAmount)
    };

    // Add date if provided
    if (effectiveDate && effectiveDate !== '') {
        payload.effective_date = effectiveDate;
    }

    try {
        const response = await fetch(`${API_BASE_URL}/calculate`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'X-API-Key': apiKey
            },
            body: JSON.stringify(payload)
        });

        const data = await response.json();

        if (!response.ok || !data.success) {
            const errorMsg = data.error || data.message || `API error: ${response.status}`;
            throw new Error(errorMsg);
        }

        // Extract Base Tax and Stamping Fee from breakdown
        const breakdown = data.breakdown || {};
        const baseTax = breakdown.base_tax?.amount || 0;
        const stampingFee = breakdown.stamping_fee?.amount || 0;

        // Build result array
        const result = [
            ['Base Tax', baseTax],
            ['Stamping Fee', stampingFee]
        ];

        // Add fallback notice if present
        if (data.fallback_reason) {
            result.push(['⚠️ Notice', data.fallback_reason]);
            result.push(['Rates From', data.rates_from || 'current']);
        }

        return result;

    } catch (error) {
        throw new Error(error.message);
    }
}

/**
 * Get rates for a state and return all 9 fee fields
 * @private
 */
async function getRate(apiKey, stateCode, effectiveDate) {
    // Build URL with query parameters
    let url = `${API_BASE_URL}/historical-rates?state=${encodeURIComponent(stateCode.toString().trim())}`;

    // Add date parameter if provided
    if (effectiveDate && effectiveDate !== '') {
        url += `&date=${encodeURIComponent(effectiveDate)}`;
    }

    try {
        const response = await fetch(url, {
            method: 'GET',
            headers: {
                'X-API-Key': apiKey
            }
        });

        const data = await response.json();

        if (!response.ok || !data.success) {
            const errorMsg = data.error || data.message || `API error: ${response.status}`;
            throw new Error(errorMsg);
        }

        // Extract rate information
        const rate = data.rate || {};

        // Build result array with 9 fee fields
        const result = [
            ['tax_rate', rate.tax_rate || ''],
            ['stamping_fee', rate.stamping_fee || ''],
            ['filing_fee', rate.filing_fee || ''],
            ['service_fee', rate.service_fee || ''],
            ['surcharge', rate.surcharge || ''],
            ['regulatory_fee', rate.regulatory_fee || ''],
            ['fire_marshal_tax', rate.fire_marshal_tax || ''],
            ['slas_clearinghouse_fee', rate.slas_clearinghouse_fee || ''],
            ['flat_fee', rate.flat_fee || '']
        ];

        // Add fallback notice if present
        if (data.fallback_reason) {
            result.push(['⚠️ Notice', data.fallback_reason]);
            result.push(['Rates From', data.rates_from || 'current']);
        }

        return result;

    } catch (error) {
        throw new Error(error.message);
    }
}

// ============================================================================
// FUNCTION REGISTRATION
// ============================================================================

// Register function with Office if available
if (typeof CustomFunctions !== 'undefined') {
    CustomFunctions.associate('SLAPI', slapi);
}

// Export for testing
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        slapi,
        getApiKey,
        setApiKey
    };
}
