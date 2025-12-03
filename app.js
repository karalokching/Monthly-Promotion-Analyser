// Global state
let promotionData = [];
let rawData = [];
let storeData = [];
let baselineData = [];
let promotionSKUs = new Set();
let promotionSKUMap = new Map(); // Maps promotion ID to its SKUs
let memberChart = null;
let storeChart = null;

// Column mapping for flexible data formats
const columnMappings = {
    txDate: ['tx date', 'transaction date', 'date', 'txdate', 'xf_txdate'],
    promotionId: ['promotion id', 'promotionid', 'promo id', 'promoid'],
    promotionDesc: ['promotion desci', 'promotion description', 'promo desc', 'description'],
    storeCode: ['store code', 'storecode', 'store', 'xf_storecode'],
    vipCode: ['vip code', 'vipcode', 'customer id', 'customerid', 'xf_vipcode'],
    docNo: ['doc no', 'docno', 'document no', 'xf_docno'],
    pluStyle: ['plu style', 'style', 'sku', 'product code', 'xf_plu'],
    itemDesc: ['item description', 'item desc', 'product name', 'xf_desci'],
    brand: ['brand', 'brandlevel'],
    animalType: ['animal type', 'animaltype', 'pet type'],
    productGroup: ['product group', 'productgroup', 'xf_group0'],
    productClass: ['product class', 'productclass', 'xf_group1'],
    productCategory: ['product category', 'category', 'xf_group2'],
    productSubCategory: ['product sub-category', 'subcategory', 'sub category'],
    qtySold: ['qty sold', 'quantity', 'qty', 'xf_qtysold'],
    amtSold: ['amt sold', 'amount', 'revenue', 'sales', 'xf_amtsold'],
    promLess: ['prom less', 'discount', 'promotion discount'],
    ttlSellPrice: ['ttl sell price', 'sell price', 'selling price'],
    ttlOrgPrice: ['ttl org price', 'original price', 'org price']
};

// Find column name in data
function findColumn(headers, mappings) {
    const lowerHeaders = headers.map(h => h.toLowerCase().trim());
    for (let mapping of mappings) {
        const index = lowerHeaders.indexOf(mapping.toLowerCase());
        if (index !== -1) return headers[index];
    }
    return null;
}

// Event listeners
document.getElementById('fileInput').addEventListener('change', handleFileSelect);
document.getElementById('processBtn').addEventListener('click', processData);
document.getElementById('baselineInput').addEventListener('change', handleBaselineSelect);
document.getElementById('processBaselineBtn').addEventListener('click', processBaseline);
document.getElementById('calculateExtraSalesBtn').addEventListener('click', calculateExtraSales);
document.getElementById('exportBtn').addEventListener('click', exportToExcel);
document.getElementById('searchInput').addEventListener('input', filterTable);
document.getElementById('promotionFilter').addEventListener('change', updateChartForPromotion);

function handleFileSelect(event) {
    const file = event.target.files[0];
    if (file) {
        document.getElementById('processBtn').disabled = false;
        document.getElementById('fileStatus').textContent = `Selected: ${file.name}`;
    }
}

function handleBaselineSelect(event) {
    const file = event.target.files[0];
    if (file) {
        document.getElementById('processBaselineBtn').disabled = false;
        document.getElementById('baselineStatus').textContent = `Selected: ${file.name}`;
    }
}

async function processData() {
    const fileInput = document.getElementById('fileInput');
    const file = fileInput.files[0];
    
    if (!file) {
        alert('Please select a file first');
        return;
    }

    document.getElementById('fileStatus').textContent = 'Processing...';
    
    try {
        const data = await readFile(file);
        rawData = data;
        analyzePromotions(data);
        displayResults();
        document.getElementById('resultsSection').style.display = 'block';
        document.getElementById('fileStatus').textContent = 'Processing complete!';
    } catch (error) {
        console.error('Error processing file:', error);
        document.getElementById('fileStatus').textContent = 'Error processing file. Please check the format.';
    }
}

function readFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array', raw: true });
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                const jsonData = XLSX.utils.sheet_to_json(firstSheet, { raw: false, defval: '' });
                resolve(jsonData);
            } catch (error) {
                console.error('File reading error:', error);
                reject(error);
            }
        };
        
        reader.onerror = () => {
            console.error('FileReader error:', reader.error);
            reject(reader.error);
        };
        reader.readAsArrayBuffer(file);
    });
}

function analyzePromotions(data) {
    if (data.length === 0) return;
    
    // Get column names
    const headers = Object.keys(data[0]);
    const cols = {
        txDate: findColumn(headers, columnMappings.txDate),
        promotionId: findColumn(headers, columnMappings.promotionId),
        promotionDesc: findColumn(headers, columnMappings.promotionDesc),
        storeCode: findColumn(headers, columnMappings.storeCode),
        vipCode: findColumn(headers, columnMappings.vipCode),
        pluStyle: findColumn(headers, columnMappings.pluStyle),
        qtySold: findColumn(headers, columnMappings.qtySold),
        amtSold: findColumn(headers, columnMappings.amtSold),
        promLess: findColumn(headers, columnMappings.promLess),
        ttlOrgPrice: findColumn(headers, columnMappings.ttlOrgPrice)
    };

    // Group by promotion
    const promotionMap = new Map();
    const storeMap = new Map();
    
    data.forEach(row => {
        const promoId = row[cols.promotionId];
        const promoDesc = row[cols.promotionDesc];
        const storeCode = row[cols.storeCode];
        const vipCode = row[cols.vipCode];
        const pluStyle = row[cols.pluStyle];
        const qtySold = parseFloat(row[cols.qtySold]) || 0;
        const amtSold = parseFloat(row[cols.amtSold]) || 0;
        const promLess = parseFloat(row[cols.promLess]) || 0;
        const orgPrice = parseFloat(row[cols.ttlOrgPrice]) || 0;
        
        if (!promoId) return;
        
        // Track SKUs in promotions
        if (pluStyle) {
            promotionSKUs.add(pluStyle);
            
            // Track SKUs per promotion
            if (!promotionSKUMap.has(promoId)) {
                promotionSKUMap.set(promoId, new Set());
            }
            promotionSKUMap.get(promoId).add(pluStyle);
        }
        
        // Track promotion data
        if (!promotionMap.has(promoId)) {
            promotionMap.set(promoId, {
                promotionId: promoId,
                description: promoDesc || '',
                newMembers: new Set(),
                existingMembers: new Set(),
                qtySold: 0,
                revenue: 0,
                discount: 0,
                originalPrice: 0
            });
        }
        
        const promo = promotionMap.get(promoId);
        
        // Check if VIP code is blank/empty (new member)
        const isNewMember = !vipCode || vipCode.toString().trim() === '';
        
        if (isNewMember) {
            promo.newMembers.add('NEW_' + Math.random()); // Each blank is a unique new member
        } else {
            promo.existingMembers.add(vipCode);
        }
        
        promo.qtySold += qtySold;
        promo.revenue += amtSold;
        promo.discount += promLess;
        promo.originalPrice += orgPrice;
        
        // Track store performance per promotion
        const storeKey = `${promoId}|${storeCode}`;
        if (!storeMap.has(storeKey)) {
            storeMap.set(storeKey, {
                promotionId: promoId,
                storeCode: storeCode || 'Unknown',
                usage: 0,
                revenue: 0,
                qtySold: 0
            });
        }
        
        const store = storeMap.get(storeKey);
        store.usage += 1;
        store.revenue += amtSold;
        store.qtySold += qtySold;
    });
    
    // Store data for later use
    storeData = Array.from(storeMap.values());
    
    // Convert to array and calculate metrics
    promotionData = Array.from(promotionMap.values()).map(promo => ({
        ...promo,
        newMemberCount: promo.newMembers.size,
        existingMemberCount: promo.existingMembers.size,
        totalCustomers: promo.newMembers.size + promo.existingMembers.size,
        discountPercent: promo.originalPrice > 0 ? (promo.discount / promo.originalPrice * 100) : 0
    }));
    
    // Sort by revenue descending
    promotionData.sort((a, b) => b.revenue - a.revenue);
}

function displayResults() {
    // Update summary cards
    const totalPromotions = promotionData.length;
    const totalTransactions = rawData.length;
    const totalRevenue = promotionData.reduce((sum, p) => sum + p.revenue, 0);
    const totalDiscount = promotionData.reduce((sum, p) => sum + p.discount, 0);
    
    document.getElementById('totalPromotions').textContent = totalPromotions;
    document.getElementById('totalTransactions').textContent = totalTransactions.toLocaleString();
    document.getElementById('totalRevenue').textContent = '$' + totalRevenue.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2});
    document.getElementById('totalDiscount').textContent = '$' + totalDiscount.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2});
    
    // Populate promotion filter dropdown
    populatePromotionFilter();
    
    // Create member chart
    createMemberChart();
    
    // Create store chart
    createStoreChart();
    
    // Populate table
    populateTable(promotionData);
}

function populatePromotionFilter() {
    const select = document.getElementById('promotionFilter');
    select.innerHTML = '<option value="all">All Promotions</option>';
    
    promotionData.forEach(promo => {
        const option = document.createElement('option');
        option.value = promo.promotionId;
        option.textContent = `${promo.promotionId} - ${promo.description.substring(0, 50)}${promo.description.length > 50 ? '...' : ''}`;
        select.appendChild(option);
    });
}

function updateChartForPromotion() {
    const selectedPromoId = document.getElementById('promotionFilter').value;
    createMemberChart(selectedPromoId);
    createStoreChart(selectedPromoId);
}

function createMemberChart(promoId = 'all') {
    let totalNew, totalExisting, selectedPromo;
    
    if (promoId === 'all') {
        totalNew = promotionData.reduce((sum, p) => sum + p.newMemberCount, 0);
        totalExisting = promotionData.reduce((sum, p) => sum + p.existingMemberCount, 0);
    } else {
        selectedPromo = promotionData.find(p => p.promotionId === promoId);
        if (selectedPromo) {
            totalNew = selectedPromo.newMemberCount;
            totalExisting = selectedPromo.existingMemberCount;
        } else {
            totalNew = 0;
            totalExisting = 0;
        }
    }
    
    const ctx = document.getElementById('memberChart').getContext('2d');
    
    if (memberChart) {
        memberChart.destroy();
    }
    
    memberChart = new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels: ['New Members', 'Existing Members'],
            datasets: [{
                data: [totalNew, totalExisting],
                backgroundColor: ['#667eea', '#764ba2'],
                borderWidth: 2,
                borderColor: '#fff'
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            plugins: {
                legend: {
                    position: 'bottom'
                },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            const label = context.label || '';
                            const value = context.parsed || 0;
                            const total = context.dataset.data.reduce((a, b) => a + b, 0);
                            const percentage = total > 0 ? ((value / total) * 100).toFixed(1) : 0;
                            return `${label}: ${value.toLocaleString()} (${percentage}%)`;
                        }
                    }
                }
            }
        }
    });
    
    // Update chart stats
    updateChartStats(totalNew, totalExisting, selectedPromo);
}

function createStoreChart(promoId = 'all') {
    let storePerformance;
    
    if (promoId === 'all') {
        // Aggregate all stores across all promotions
        const storeAgg = new Map();
        storeData.forEach(store => {
            if (!storeAgg.has(store.storeCode)) {
                storeAgg.set(store.storeCode, {
                    storeCode: store.storeCode,
                    usage: 0,
                    revenue: 0,
                    qtySold: 0
                });
            }
            const agg = storeAgg.get(store.storeCode);
            agg.usage += store.usage;
            agg.revenue += store.revenue;
            agg.qtySold += store.qtySold;
        });
        storePerformance = Array.from(storeAgg.values());
    } else {
        // Filter stores for selected promotion
        storePerformance = storeData.filter(s => s.promotionId === promoId);
    }
    
    // Sort by revenue and take top 10
    storePerformance.sort((a, b) => b.revenue - a.revenue);
    const topStores = storePerformance.slice(0, 10);
    
    const ctx = document.getElementById('storeChart').getContext('2d');
    
    if (storeChart) {
        storeChart.destroy();
    }
    
    storeChart = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: topStores.map(s => s.storeCode),
            datasets: [{
                label: 'Revenue ($)',
                data: topStores.map(s => s.revenue),
                backgroundColor: '#667eea',
                borderColor: '#5568d3',
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            plugins: {
                legend: {
                    display: false
                },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            const store = topStores[context.dataIndex];
                            return [
                                `Revenue: $${store.revenue.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})}`,
                                `Usage: ${store.usage} transactions`,
                                `Qty Sold: ${store.qtySold.toLocaleString()}`
                            ];
                        }
                    }
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: {
                        callback: function(value) {
                            return '$' + value.toLocaleString();
                        }
                    }
                }
            }
        }
    });
    
    // Update store stats
    updateStoreStats(topStores);
}

function updateStoreStats(stores) {
    const statsDiv = document.getElementById('storeStats');
    
    if (stores.length === 0) {
        statsDiv.innerHTML = '<div class="chart-stat-item"><div class="chart-stat-label">No store data available</div></div>';
        return;
    }
    
    const topStore = stores[0];
    const totalRevenue = stores.reduce((sum, s) => sum + s.revenue, 0);
    const totalUsage = stores.reduce((sum, s) => sum + s.usage, 0);
    
    const statsHTML = `
        <div class="chart-stat-item">
            <div class="chart-stat-label">Top Store</div>
            <div class="chart-stat-value">${topStore.storeCode}</div>
        </div>
        <div class="chart-stat-item">
            <div class="chart-stat-label">Top Store Revenue</div>
            <div class="chart-stat-value">$${topStore.revenue.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})}</div>
        </div>
        <div class="chart-stat-item">
            <div class="chart-stat-label">Top Store Usage</div>
            <div class="chart-stat-value">${topStore.usage} txns</div>
        </div>
        <div class="chart-stat-item">
            <div class="chart-stat-label">Total Stores</div>
            <div class="chart-stat-value">${stores.length}</div>
        </div>
    `;
    
    statsDiv.innerHTML = statsHTML;
}

function updateChartStats(newMembers, existingMembers, promo) {
    const statsDiv = document.getElementById('chartStats');
    const total = newMembers + existingMembers;
    const newPercent = total > 0 ? ((newMembers / total) * 100).toFixed(1) : 0;
    const existingPercent = total > 0 ? ((existingMembers / total) * 100).toFixed(1) : 0;
    
    let statsHTML = `
        <div class="chart-stat-item">
            <div class="chart-stat-label">New Members</div>
            <div class="chart-stat-value">${newMembers} (${newPercent}%)</div>
        </div>
        <div class="chart-stat-item">
            <div class="chart-stat-label">Existing Members</div>
            <div class="chart-stat-value">${existingMembers} (${existingPercent}%)</div>
        </div>
        <div class="chart-stat-item">
            <div class="chart-stat-label">Total Customers</div>
            <div class="chart-stat-value">${total}</div>
        </div>
    `;
    
    if (promo) {
        statsHTML += `
            <div class="chart-stat-item">
                <div class="chart-stat-label">Revenue</div>
                <div class="chart-stat-value">$${promo.revenue.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})}</div>
            </div>
            <div class="chart-stat-item">
                <div class="chart-stat-label">Qty Sold</div>
                <div class="chart-stat-value">${promo.qtySold.toLocaleString()}</div>
            </div>
            <div class="chart-stat-item">
                <div class="chart-stat-label">Discount</div>
                <div class="chart-stat-value">$${promo.discount.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})}</div>
            </div>
        `;
    }
    
    statsDiv.innerHTML = statsHTML;
}

function populateTable(data) {
    const tbody = document.getElementById('promotionTableBody');
    tbody.innerHTML = '';
    
    data.forEach(promo => {
        const row = tbody.insertRow();
        row.innerHTML = `
            <td>${promo.promotionId}</td>
            <td>${promo.description}</td>
            <td class="number">${promo.newMemberCount}</td>
            <td class="number">${promo.existingMemberCount}</td>
            <td class="number">${promo.totalCustomers}</td>
            <td class="number">${promo.qtySold.toLocaleString()}</td>
            <td class="number">$${promo.revenue.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td>
            <td class="number">$${promo.discount.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td>
            <td class="number">${promo.discountPercent.toFixed(1)}%</td>
        `;
    });
}

function filterTable() {
    const searchTerm = document.getElementById('searchInput').value.toLowerCase();
    const filtered = promotionData.filter(promo => 
        promo.promotionId.toLowerCase().includes(searchTerm) ||
        promo.description.toLowerCase().includes(searchTerm)
    );
    populateTable(filtered);
}

async function processBaseline() {
    const fileInput = document.getElementById('baselineInput');
    const file = fileInput.files[0];
    
    if (!file) {
        alert('Please select a baseline file first');
        return;
    }

    document.getElementById('baselineStatus').textContent = 'Processing baseline data...';
    
    try {
        const data = await readFile(file);
        
        if (!data || data.length === 0) {
            throw new Error('No data found in file');
        }
        
        baselineData = data;
        
        console.log('Baseline data loaded:', data.length, 'records');
        console.log('Sample record:', data[0]);
        
        document.getElementById('baselineConfig').style.display = 'block';
        document.getElementById('baselineStatus').textContent = `Baseline data loaded: ${data.length} records`;
        
        // Auto-populate date range if possible
        const headers = Object.keys(data[0]);
        const txDateCol = findColumn(headers, columnMappings.txDate);
        
        console.log('Found date column:', txDateCol);
        
        if (txDateCol) {
            const dates = data.map(row => parseDate(row[txDateCol])).filter(d => d);
            console.log('Parsed dates:', dates.length);
            
            if (dates.length > 0) {
                const minDate = new Date(Math.min(...dates));
                const maxDate = new Date(Math.max(...dates));
                document.getElementById('baselineStartDate').value = minDate.toISOString().split('T')[0];
                document.getElementById('baselineEndDate').value = maxDate.toISOString().split('T')[0];
                
                const daysDiff = Math.ceil((maxDate - minDate) / (1000 * 60 * 60 * 24)) + 1;
                document.getElementById('baselineInfo').textContent = `Baseline period: ${daysDiff} days (${minDate.toLocaleDateString()} - ${maxDate.toLocaleDateString()})`;
            }
        } else {
            document.getElementById('baselineInfo').textContent = 'Date column not found. Available columns: ' + headers.join(', ');
        }
    } catch (error) {
        console.error('Error processing baseline file:', error);
        document.getElementById('baselineStatus').textContent = `Error: ${error.message}. Check browser console for details.`;
    }
}

function parseDate(dateStr) {
    if (!dateStr) return null;
    
    const str = dateStr.toString().trim();
    if (!str) return null;
    
    // Try standard Date parsing first
    let date = new Date(str);
    if (!isNaN(date.getTime()) && date.getFullYear() > 1900) return date;
    
    // Try MM/DD/YYYY HH:MM:SS format (common in exports)
    const parts = str.split(/[\s\/\-:]/);
    if (parts.length >= 3) {
        // MM/DD/YYYY format
        const month = parseInt(parts[0]);
        const day = parseInt(parts[1]);
        const year = parseInt(parts[2]);
        
        if (month >= 1 && month <= 12 && day >= 1 && day <= 31 && year > 1900) {
            date = new Date(year, month - 1, day);
            if (!isNaN(date.getTime())) return date;
        }
    }
    
    // Try DD/MM/YYYY format
    if (parts.length >= 3) {
        const day = parseInt(parts[0]);
        const month = parseInt(parts[1]);
        const year = parseInt(parts[2]);
        
        if (month >= 1 && month <= 12 && day >= 1 && day <= 31 && year > 1900) {
            date = new Date(year, month - 1, day);
            if (!isNaN(date.getTime())) return date;
        }
    }
    
    return null;
}

function calculateExtraSales() {
    if (baselineData.length === 0) {
        alert('Please upload and process baseline data first');
        return;
    }
    
    const startDate = new Date(document.getElementById('baselineStartDate').value);
    const endDate = new Date(document.getElementById('baselineEndDate').value);
    
    if (!startDate || !endDate || isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
        alert('Please select valid baseline start and end dates');
        return;
    }
    
    // Get baseline data headers
    const headers = Object.keys(baselineData[0]);
    const cols = {
        txDate: findColumn(headers, columnMappings.txDate),
        pluStyle: findColumn(headers, columnMappings.pluStyle),
        qtySold: findColumn(headers, columnMappings.qtySold),
        amtSold: findColumn(headers, columnMappings.amtSold)
    };
    
    // Calculate baseline revenue for promotion SKUs
    let baselineRevenue = 0;
    let baselineDays = 0;
    
    baselineData.forEach(row => {
        const txDate = parseDate(row[cols.txDate]);
        const pluStyle = row[cols.pluStyle];
        const amtSold = parseFloat(row[cols.amtSold]) || 0;
        
        if (txDate && txDate >= startDate && txDate <= endDate && promotionSKUs.has(pluStyle)) {
            baselineRevenue += amtSold;
        }
    });
    
    baselineDays = Math.ceil((endDate - startDate) / (1000 * 60 * 60 * 24)) + 1;
    
    // Get promotion period
    const promoHeaders = Object.keys(rawData[0]);
    const promoTxDateCol = findColumn(promoHeaders, columnMappings.txDate);
    
    let promoDates = [];
    if (promoTxDateCol) {
        promoDates = rawData.map(row => parseDate(row[promoTxDateCol])).filter(d => d);
    }
    
    let promoDays = 1;
    if (promoDates.length > 0) {
        const minPromoDate = new Date(Math.min(...promoDates));
        const maxPromoDate = new Date(Math.max(...promoDates));
        promoDays = Math.ceil((maxPromoDate - minPromoDate) / (1000 * 60 * 60 * 24)) + 1;
    }
    
    // Calculate daily baseline and scale to promotion period
    const dailyBaseline = baselineDays > 0 ? baselineRevenue / baselineDays : 0;
    const scaledBaseline = dailyBaseline * promoDays;
    
    // Get total promotion revenue
    const promoRevenue = promotionData.reduce((sum, p) => sum + p.revenue, 0);
    
    // Calculate extra sales
    const extraSales = promoRevenue - scaledBaseline;
    const uplift = scaledBaseline > 0 ? ((extraSales / scaledBaseline) * 100) : 0;
    
    // Calculate extra sales by promotion
    const extraSalesByPromo = calculateExtraSalesByPromotion(startDate, endDate, baselineDays, promoDays, dailyBaseline, cols);
    
    // Display results
    document.getElementById('extraSalesSection').style.display = 'block';
    document.getElementById('extraSales').textContent = '$' + extraSales.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2});
    document.getElementById('promoRevenue').textContent = '$' + promoRevenue.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2});
    document.getElementById('baselineRevenue').textContent = '$' + scaledBaseline.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2});
    document.getElementById('upliftPercent').textContent = uplift.toFixed(1) + '%';
    
    // Update baseline info
    document.getElementById('baselineInfo').innerHTML = `
        <strong>Baseline Period:</strong> ${baselineDays} days (${startDate.toLocaleDateString()} - ${endDate.toLocaleDateString()})<br>
        <strong>Promotion Period:</strong> ${promoDays} days<br>
        <strong>Daily Baseline:</strong> $${dailyBaseline.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})}<br>
        <strong>Scaled Baseline:</strong> $${scaledBaseline.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})} (${promoDays} days)
    `;
    
    // Populate extra sales by promotion table
    populateExtraSalesTable(extraSalesByPromo);
}

function calculateExtraSalesByPromotion(startDate, endDate, baselineDays, promoDays, dailyBaseline, baselineCols) {
    const promoExtraSales = [];
    
    promotionData.forEach(promo => {
        const promoSKUs = promotionSKUMap.get(promo.promotionId) || new Set();
        
        // Calculate baseline revenue and quantity for this promotion's SKUs
        let promoBaselineRevenue = 0;
        let promoBaselineQty = 0;
        
        baselineData.forEach(row => {
            const txDate = parseDate(row[baselineCols.txDate]);
            const pluStyle = row[baselineCols.pluStyle];
            const amtSold = parseFloat(row[baselineCols.amtSold]) || 0;
            const qtySold = parseFloat(row[baselineCols.qtySold]) || 0;
            
            if (txDate && txDate >= startDate && txDate <= endDate && promoSKUs.has(pluStyle)) {
                promoBaselineRevenue += amtSold;
                promoBaselineQty += qtySold;
            }
        });
        
        // Scale baseline to promotion period
        const promoScaledBaseline = baselineDays > 0 ? (promoBaselineRevenue / baselineDays) * promoDays : 0;
        const promoScaledBaselineQty = baselineDays > 0 ? (promoBaselineQty / baselineDays) * promoDays : 0;
        
        // Calculate extra sales for this promotion
        const promoExtraSalesValue = promo.revenue - promoScaledBaseline;
        const promoExtraQty = promo.qtySold - promoScaledBaselineQty;
        const promoUplift = promoScaledBaseline > 0 ? ((promoExtraSalesValue / promoScaledBaseline) * 100) : 0;
        const roi = promo.discount > 0 ? ((promoExtraSalesValue / promo.discount) * 100) : 0;
        
        promoExtraSales.push({
            promotionId: promo.promotionId,
            description: promo.description,
            qtySold: promo.qtySold,
            baselineQty: promoScaledBaselineQty,
            extraQty: promoExtraQty,
            revenue: promo.revenue,
            baselineRevenue: promoScaledBaseline,
            extraSales: promoExtraSalesValue,
            uplift: promoUplift,
            discount: promo.discount,
            roi: roi
        });
    });
    
    // Sort by extra sales descending
    promoExtraSales.sort((a, b) => b.extraSales - a.extraSales);
    
    return promoExtraSales;
}

function populateExtraSalesTable(data) {
    const tbody = document.getElementById('extraSalesTableBody');
    tbody.innerHTML = '';
    
    data.forEach(promo => {
        const row = tbody.insertRow();
        const extraSalesClass = promo.extraSales >= 0 ? 'positive' : 'negative';
        const extraQtyClass = promo.extraQty >= 0 ? 'positive' : 'negative';
        
        row.innerHTML = `
            <td>${promo.promotionId}</td>
            <td>${promo.description}</td>
            <td class="number">${promo.qtySold.toLocaleString()}</td>
            <td class="number">${promo.baselineQty.toLocaleString(undefined, {minimumFractionDigits: 0, maximumFractionDigits: 0})}</td>
            <td class="number ${extraQtyClass}">${promo.extraQty.toLocaleString(undefined, {minimumFractionDigits: 0, maximumFractionDigits: 0})}</td>
            <td class="number">$${promo.revenue.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td>
            <td class="number">$${promo.baselineRevenue.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td>
            <td class="number ${extraSalesClass}">$${promo.extraSales.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td>
            <td class="number">${promo.uplift.toFixed(1)}%</td>
            <td class="number">${promo.roi.toFixed(1)}%</td>
        `;
    });
}

function exportToExcel() {
    const exportData = promotionData.map(promo => ({
        'Promotion ID': promo.promotionId,
        'Description': promo.description,
        'New Members': promo.newMemberCount,
        'Existing Members': promo.existingMemberCount,
        'Total Customers': promo.totalCustomers,
        'Qty Sold': promo.qtySold,
        'Revenue': promo.revenue,
        'Discount': promo.discount,
        'Discount %': promo.discountPercent.toFixed(2)
    }));
    
    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Promotion Review');
    
    const fileName = `promotion_review_${new Date().toISOString().split('T')[0]}.xlsx`;
    XLSX.writeFile(wb, fileName);
}
