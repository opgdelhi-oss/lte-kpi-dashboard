// sampleData.js - Realistic 14-Day LTE RAN KPI Dataset Generator
// Simulates real-world cellular network behavior: diurnal sinusoidal traffic curves,
// evening busy hours (20:00 - 22:00), weekday vs weekend variations, and realistic degradations.

(function (root, factory) {
    if (typeof define === 'function' && define.amd) {
        define([], factory);
    } else if (typeof module === 'object' && module.exports) {
        module.exports = factory();
    } else {
        root.SampleLTEData = factory();
    }
}(typeof self !== 'undefined' ? self : this, function () {

    const SITES = [
        { id: 'eNB_101', name: 'Downtown_Commercial', cluster: 'Metro_Core', cells: 3, band: 'B3 (1800MHz)', duplex: 'FDD' },
        { id: 'eNB_102', name: 'TechPark_Sector', cluster: 'HighTech_Corridor', cells: 3, band: 'B40 (2300MHz)', duplex: 'TDD' },
        { id: 'eNB_103', name: 'Suburban_Residency', cluster: 'Residential_North', cells: 3, band: 'B3 (1800MHz)', duplex: 'FDD' },
        { id: 'eNB_104', name: 'Airport_Expressway', cluster: 'Transport_Hub', cells: 2, band: 'B5 (850MHz)', duplex: 'FDD' },
        { id: 'eNB_105', name: 'CityCenter_Mall', cluster: 'Metro_Core', cells: 3, band: 'B40 (2300MHz)', duplex: 'TDD' },
        { id: 'eNB_106', name: 'University_Campus', cluster: 'Education_District', cells: 3, band: 'B3 (1800MHz)', duplex: 'FDD' }
    ];

    // Pseudo-random deterministic generator for consistent sample data
    function pseudoRandom(seed) {
        const x = Math.sin(seed++) * 10000;
        return x - Math.floor(x);
    }

    function generateDataset(daysCount = 14) {
        const data = [];
        const endDate = new Date(2026, 8, 22); // Target current local reference date
        let seed = 42;

        for (let d = daysCount - 1; d >= 0; d--) {
            const currentDate = new Date(endDate);
            currentDate.setDate(currentDate.getDate() - d);
            const dateStr = currentDate.toISOString().split('T')[0];
            const isWeekend = currentDate.getDay() === 0 || currentDate.getDay() === 6;

            for (let hour = 0; hour < 24; hour++) {
                // Diurnal curve: peaks around 12:00-14:00 (lunch) and 20:00-22:00 (evening peak), lowest at 04:00
                const hourAngle = (hour - 4) * (Math.PI / 12);
                let diurnalFactor = 0.25 + 0.75 * Math.max(0, Math.sin(hourAngle / 2));
                if (hour >= 19 && hour <= 22) diurnalFactor = 1.0; // Evening peak busy hour
                if (hour >= 2 && hour <= 5) diurnalFactor = 0.12;  // Deep sleep hours

                SITES.forEach(site => {
                    for (let c = 1; c <= site.cells; c++) {
                        seed += 1.37;
                        const rnd1 = pseudoRandom(seed);
                        const rnd2 = pseudoRandom(seed + 100);
                        const rnd3 = pseudoRandom(seed + 200);

                        const cellName = `${site.id}_${site.name}_S${c}`;
                        const cellId = parseInt(site.id.replace('eNB_', '')) * 10 + c;

                        // Site specific profile factors
                        let siteLoadFactor = site.id === 'eNB_105' ? 1.25 : (site.id === 'eNB_102' && !isWeekend ? 1.3 : 1.0);
                        let load = Math.min(0.98, diurnalFactor * siteLoadFactor * (0.8 + 0.4 * rnd1));

                        // Introduce realistic anomalies:
                        // Anomaly 1: eNB_105 Sector 2 has external interference on Day 10-12 (Drops increase, DL thrpt falls)
                        let isDegradedSector = (site.id === 'eNB_105' && c === 2 && d >= 2 && d <= 4);
                        // Anomaly 2: eNB_102 Sector 1 has severe congestion during busy hours (PRB > 88%)
                        let isCongestedSector = (site.id === 'eNB_102' && c === 1 && hour >= 19 && hour <= 21);

                        // Baseline KPIs
                        let rrcSr = 99.1 + 0.8 * rnd2;
                        let erabSr = 99.0 + 0.9 * rnd3;
                        let dropRate = 0.22 + 0.25 * rnd1;
                        let hoSr = 97.8 + 1.8 * rnd2;
                        let availability = 100.0;
                        let dlPrb = Math.min(96, Math.max(8, load * 85 + (rnd2 - 0.5) * 8));
                        let ulPrb = Math.min(85, Math.max(5, dlPrb * 0.65 + (rnd3 - 0.5) * 6));

                        // User Throughput (inversely related to PRB utilization)
                        let baseDlSpeed = site.band.includes('B40') ? 42 : 28; // TDD 2300MHz higher capacity
                        let dlThroughput = Math.max(3.2, baseDlSpeed * (1 - (dlPrb / 100) * 0.68) + (rnd1 - 0.5) * 4);
                        let ulThroughput = Math.max(1.8, (dlThroughput * 0.35) + (rnd2 - 0.5) * 2);

                        // Traffic volume in GB
                        let baseTraffic = site.band.includes('B40') ? 14 : 9;
                        let dlTrafficGb = parseFloat((baseTraffic * load * (0.9 + 0.2 * rnd3)).toFixed(2));
                        let ulTrafficGb = parseFloat((dlTrafficGb * 0.18 * (0.95 + 0.1 * rnd1)).toFixed(2));

                        // Active users
                        let avgUsers = Math.max(8, Math.round(load * 75 + (rnd1 - 0.5) * 12));
                        let maxUsers = Math.round(avgUsers * (1.35 + 0.2 * rnd2));

                        // CQI >= 7 percentage
                        let cqi7Rate = 88.5 + 8.5 * rnd1;

                        // Apply Anomaly Effects
                        if (isDegradedSector) {
                            dropRate += 2.4 + rnd2 * 1.5; // Elevated drop rate up to 3.8%
                            rrcSr -= (2.5 + rnd1 * 2);    // RRC SR drops to ~94%
                            dlThroughput *= 0.45;         // Severe throughput degradation
                            cqi7Rate -= 22;               // Poor channel quality
                        }

                        if (isCongestedSector) {
                            dlPrb = 89 + rnd1 * 8;        // Critical PRB congestion > 90%
                            dlThroughput = Math.max(2.1, 4.5 - rnd2 * 1.8); // Sluggish throughput
                            avgUsers += 40;
                            maxUsers += 65;
                        }

                        // Rare cell outage (0.2% chance)
                        if (rnd1 < 0.002) {
                            availability = 0;
                            dlTrafficGb = 0;
                            ulTrafficGb = 0;
                            avgUsers = 0;
                            maxUsers = 0;
                            rrcSr = 0;
                            erabSr = 0;
                        }

                        data.push({
                            'Date': dateStr,
                            'Hour': hour,
                            'Timestamp': `${dateStr} ${String(hour).padStart(2, '0')}:00`,
                            'Cluster': site.cluster,
                            'Site_ID': site.id,
                            'Site_Name': site.name,
                            'Cell_Name': cellName,
                            'Cell_ID': cellId,
                            'Band': site.band,
                            'Duplex': site.duplex,
                            'RRC_Setup_SR': parseFloat(rrcSr.toFixed(2)),
                            'ERAB_Setup_SR': parseFloat(erabSr.toFixed(2)),
                            'Drop_Rate': parseFloat(dropRate.toFixed(2)),
                            'Handover_SR': parseFloat(hoSr.toFixed(2)),
                            'DL_User_Throughput_Mbps': parseFloat(dlThroughput.toFixed(2)),
                            'UL_User_Throughput_Mbps': parseFloat(ulThroughput.toFixed(2)),
                            'DL_Traffic_GB': dlTrafficGb,
                            'UL_Traffic_GB': ulTrafficGb,
                            'Total_Traffic_GB': parseFloat((dlTrafficGb + ulTrafficGb).toFixed(2)),
                            'DL_PRB_Util': parseFloat(dlPrb.toFixed(1)),
                            'UL_PRB_Util': parseFloat(ulPrb.toFixed(1)),
                            'Active_Users_Avg': avgUsers,
                            'Active_Users_Max': maxUsers,
                            'Cell_Availability': parseFloat(availability.toFixed(1)),
                            'CQI_ge_7_Rate': parseFloat(cqi7Rate.toFixed(1))
                        });
                    }
                });
            }
        }

        return data;
    }

    return {
        generateDataset: generateDataset,
        sites: SITES
    };
}));
