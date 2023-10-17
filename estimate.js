let xlsxCellsConstValues = {
    'D4': 1,
    'E5': 3000,
    'E11': 5,
    'F4': 0,
}

function createCell(cellValue) {
    const cell = document.createElement('td');
    cell.textContent = cellValue;

    return cell;
}

function estimate() {
    const form = document.querySelector('form');

    const depth = Number(form.querySelector('input[name="depth"]').value);
    const tubingInnerDiameter = Number(form.querySelector('input[name="tubing-inner-diameter"]').value);
    const oilGravity = Number(form.querySelector('input[name="oil-gravity"]').value);
    const oilViscosity = Number(form.querySelector('input[name="oil-viscosity"]').value);
    const gasLiquidRatio = Number(form.querySelector('input[name="gas-liquid-ratio"]').value);
    const gasSpecificGravity = Number(form.querySelector('input[name="gas-specific-gravity"]').value);
    const flowingTubingHeadPressure = Number(form.querySelector('input[name="flowing-tubing-head-pressure"]').value);
    const flowingTubingHeadTemperature = Number(form.querySelector('input[name="flowing-tubing-head-temperature"]').value);
    const flowingTemperatureAtShoe = Number(form.querySelector('input[name="flowing-temperature-at-shoe"]').value);
    const liquidProductionRate = Number(form.querySelector('input[name="liquid-production-rate"]').value);
    const waterCut = Number(form.querySelector('input[name="water-cut"]').value);
    const interfacialTension = Number(form.querySelector('input[name="interfacial-tension"]').value);
    const waterSpecialGravity = Number(form.querySelector('input[name="water-special-gravity"]').value);

    const AC12 = -2.69851, AC13 = 0.15840954, AC14 = -0.55099756, AC15 = 0.54784917, AC16 = -0.12194578;
    const AE12 = -0.10306578, AE13 = 0.617774, AE14 = -0.632946, AE15 = 0.29598, AE16 = -0.0401;
    const AI12 = 0.91162574, AI13 = -4.82175636, AI14 = 1232.25036621, AI15 = -22253.57617, AI16 = 116174/28125;
    const E6 = 0.5, E7 = 0.8, E8 = 0.002, E9 = 90, E10 = 0.709, E12 = 80, E13 = 90, E14 = 300, E15 = 0, E16 = 0.03, E17 = 1.076;
    const H5 = xlsxCellsConstValues['D4'] * depth + xlsxCellsConstValues['F4'] * xlsxCellsConstValues['E5'] * 3.28, 
        H12 = xlsxCellsConstValues['D4'] * flowingTubingHeadTemperature + xlsxCellsConstValues['F4'] * (32 + 9/5 * E12), 
        H13 = xlsxCellsConstValues['D4'] * flowingTemperatureAtShoe + xlsxCellsConstValues['F4'] * (32 + 9/5 * E13);
    const Q3 = -2.462, Q4 = 2.97, Q5 = -2.862 / 10, Q6 = 8.054 / 1000, Q7 = 2.808, Q8 = -3.498, Q9 = 3.603 / 10, Q10 = -1.044 / 100, 
        Q11 = -7.933 / 10, Q12 = 1.396, Q13 = -1.491 / 10, Q14 = 4.41 / 1000, Q15 = 8.393 / 100, Q16 = -1.864 / 10, Q17 = 2.033 / 100, Q18 = -6.095 / 1E4;
    
    const TValues = [0.0124657507677914, 0.0159076685616638, 0.0212047356079062, 0.0261421255525684, 0.0313875289401125, 0.0368761597287921, 0.042652688368104,
        0.0482699885874431, 0.0546101307401282, 0.0609352698907551, 0.0673823906190183, 0.073915217376841, 0.0805067379019854, 0.0870695279626679, 0.0936235439140799, 
        0.100100467912485, 0.106461762781915, 0.112672075238037, 0.118700330692457, 0.123459799343456, 0.12995642991769, 0.135613749610385, 0.140756592358395, 0.145640732592793,
        0.150265071038872, 0.154632361836925, 0.158748403128501, 0.162621336355577, 0.16626097025529, 0.169678157950242]

    const resultTable = document.getElementById('resultTable');
    document.getElementById('resultContainer').classList.remove('d-none');
    resultTable.innerHTML = '';
        

    let startDepth = 0;
    const depthSummand = (H5 - startDepth) / 29;

    let depthFoots = [], depthMeters = [], psiaValues = [], mPaValues = [];
    
    depthFoots.push(startDepth);
    depthMeters.push(startDepth);

    while (startDepth <= depth) {
        if (startDepth > depth) {
            break;
        }

        const currentDepthFoot = Math.round(startDepth + depthSummand);
        const currentDepthMeter = Math.round(currentDepthFoot / 3.28);

        depthFoots.push(currentDepthFoot);
        depthMeters.push(currentDepthMeter);

        startDepth = currentDepthFoot;
    } 

    let startPressure = xlsxCellsConstValues['D4'] * flowingTubingHeadPressure 
        + xlsxCellsConstValues['F4'] * xlsxCellsConstValues['E11'] * 10 * 14.7;

    psiaValues.push(startPressure);
    mPaValues.push((startPressure / 14.7 / 10).toFixed(2));

    for (let index = 1; index < depthFoots.length; index++) {
        const previousPressureValue = psiaValues[index - 1];
        const currentDepthFoot = depthFoots[index];

        let I21;

        if (index == 1) {
            I21 = H12;
        } else {
            I21 = H12 + (H13 - H12) / H5 * currentDepthFoot;
        }

        const H14 = xlsxCellsConstValues['D4'] * liquidProductionRate + xlsxCellsConstValues['F4'] * E14 / 0.159;
        const H6 = xlsxCellsConstValues['D4'] * tubingInnerDiameter + xlsxCellsConstValues['F4'] * E6 * 100 / 2.54;
        const K6 = 3.14 / 4 * Math.pow(H6 / 12, 2);
        const N14 = H14 * 5.615 / 86400 / K6;
        const H15 = xlsxCellsConstValues['D4'] * waterCut + xlsxCellsConstValues['F4'] * E15
        const K15 = H14 / 100 * H15;
        const H7 = xlsxCellsConstValues['D4'] * oilGravity + xlsxCellsConstValues['F4'] * (141.5 / E7 - 131.5);
        const K7 = 141.5 / (131.5 + H7);
        const H17 = xlsxCellsConstValues['D4'] * waterSpecialGravity + xlsxCellsConstValues['F4'] * E17
        const K14 = ((H14 - K15) * K7 + K15 * H17) / H14;
        const H16 = xlsxCellsConstValues['D4'] * interfacialTension + xlsxCellsConstValues['F4'] * E16 * 1000;
        const Y21 = 1.938 * N14 * Math.pow(62.4 * K14 / H16, 0.25);
        const H9 = xlsxCellsConstValues['D4'] * gasLiquidRatio + xlsxCellsConstValues['F4'] * E9 * 5.615;
        const K9 = H9 * H14;
        const H10 = xlsxCellsConstValues['D4'] * gasSpecificGravity + xlsxCellsConstValues['F4'] * E10;
        const S1 = 169 + 314 * H10;
        const S2 = 708.75 - 57.7 * H10;
        const K21 = (I21 + 460) / S1;
        const O21 = 1 / K21;
        const P21 = 0.06125 * O21 * Math.exp(-1.2 * Math.pow(1 - O21, 2)); 
        const J21 = previousPressureValue / S2;
        const V21 = P21 * J21 / TValues[index - 1];
        const W21 = 1 / K6 * K9 * V21 * (460 + I21) / (460 + 60) * (14.7 / previousPressureValue) / 86400
        const Z21 = 1.938 * W21 * Math.pow(62.4 * K14 / H16, 0.25);
        const H8 =  xlsxCellsConstValues['D4'] * oilViscosity + xlsxCellsConstValues['F4'] * E8 * 1000;
        const M17 = (H8 * (H14 - K15) + 0.5 * K15) / H14;
        const AB21 = 0.15726 * M17 * (1 / Math.pow(62.4 * K14 * Math.pow(H16, 3), 0.25));
        const AC21 = Math.pow(10, (AC12 + AC13 * (Math.log10(AB21) + 3) + AC14 * Math.pow(Math.log10(AB21) + 3, 2) + 
            AC15 * Math.pow(Math.log10(AB21) + 3, 3) + AC16 * Math.pow(Math.log10(AB21) + 3, 4)));
        const AA21 = 120.872 * H6 / 12 * Math.sqrt(62.4 * K14 / H16);
        const AD21 = Y21 / Math.pow(Z21, 0.575) * Math.pow(previousPressureValue / 14.7, 0.1) * AC21 / AA21;
        const AE21 = AE12 + AE13 * (Math.log10(AD21) + 6) + AE14 * Math.pow(Math.log10(AD21) + 6, 2) + 
            AE15 * Math.pow(Math.log10(AD21) + 6, 3) + AE16 * Math.pow(Math.log10(AD21) + 6, 4);
        const AF21 = Z21 * Math.pow(AB21, 0.38) / Math.pow(AA21, 2.14);
        const AG21 = (AF21 - 0.012) / Math.abs(AF21 - 0.012);
        const AH21 = (1 - AG21) / 2 * 0.012 + (1 + AG21) / 2 * AF21;
        const AI21 = AI12 + AI13 * AH21 + AI14 * Math.pow(AH21, 2) + AI15 * Math.pow(AH21, 3) + AI16 * Math.pow(AH21, 4);
        const AJ21 = AE21 * AI21;
        const AM21 = 28.97 * H10 * previousPressureValue / V21 / 10.73 / (460 + I21);
        const AN21 = AJ21 * K14 * 62.4 + (1 - AJ21) * AM21;
        const M16 = K14 * 62.4 * H14 * 5.615 + 0.0765 * H10 * K9;
        const M21 = (1.709 / 1E5 - 2.062 / 1E6 * H10) * I21 + 8.188 / 1000 - 6.15 / 1000 * Math.log(H10);
        const L21 = Q3 + Q4 * J21 + Q5 * Math.pow(J21, 2) + Q6 *  Math.pow(J21, 3) + 
            K21 * (Q7 + Q8 * J21 + Q9 * Math.pow(J21, 2) + Q10 * Math.pow(J21, 3)) + 
                Math.pow(K21, 2) * (Q11 + Q12 * J21 + Q13 * Math.pow(J21, 2) + Q14 * Math.pow(J21, 3)) + 
                    Math.pow(K21, 3) * (Q15 + Q16 * J21 + Q17 * Math.pow(J21, 2) + Q18 * Math.pow(J21, 3));
        const N21 = M21 / K21 * Math.exp(L21);
        const AK21 = 0.022 * M16 / (H6 * Math.pow(M17, AJ21) * Math.pow(N21, 1 - AJ21));
        const N6 = 0.0006;
        const AL21 = 1 / Math.pow(-4 * Math.log10(N6 / 3.7065 - 5.0452 / AK21 
            * Math.log10(Math.pow(N6, 1.1098) / 2.8257 + Math.pow(7.149 / AK21, 0.8981))), 2)
        const AO21 = 1 / 144 * (AN21 + AL21 * Math.pow(M16, 2) / 7.413 / 1E10 / Math.pow(H6 / 12, 5) / AN21);

        const psiaPressure = Math.round(previousPressureValue + AO21 * (currentDepthFoot - depthFoots[index - 1]));
        const mPaPressure = (psiaPressure / 14.7 / 10).toFixed(2);

        psiaValues.push(psiaPressure);
        mPaValues.push(mPaPressure);
    }

    for (let index = 0; index < depthFoots.length; index++) {
        const newRow = document.createElement('tr');
        newRow.className = 'text-center';

        newRow.appendChild(createCell(depthFoots[index]));
        newRow.appendChild(createCell(depthMeters[index]));
        newRow.appendChild(createCell(psiaValues[index]));
        newRow.appendChild(createCell(mPaValues[index]));

        resultTable.appendChild(newRow);
    }

    drawCharts(depthFoots, depthMeters, psiaValues, mPaValues);
}

function drawCharts(depthFoots, depthMeters, psiaValues, mPaValues) {
    const defaultOption = {
        scales: {
            x: {
                title: {
                    display: true,
                    text: 'Глубина'
                }
            },
            y: {
                title: {
                    display: true,
                    text: 'Давление'
                }
            }
        }
    };

    const englishData = {
        labels: depthFoots,
        datasets: [{
            label: 'Давление (psi)',
            data: psiaValues,
            backgroundColor: 'rgba(75, 192, 192, 0.2)',
            borderColor: 'rgba(75, 192, 192, 1)',
            borderWidth: 1,
        }]
    };
    
    const siData = {
        labels: depthMeters,
        datasets: [{
            label: 'Давление (кПа)',
            data: mPaValues,
            backgroundColor: 'rgba(255, 99, 132, 0.2)',
            borderColor: "rgba(255, 99, 132, 1)",
            borderWidth: 1,
        }]
    };
    

    const englishChartConfig = {
        type: 'line',
        data: englishData,
        options: defaultOption
    };

    const siChartConfig = {
        type: 'line',
        data: siData,
        options: defaultOption
    };

    new Chart(document.getElementById('englishChart'), englishChartConfig);
    new Chart(document.getElementById('siChart'), siChartConfig);
}