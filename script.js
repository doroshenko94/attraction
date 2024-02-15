window.onload = function() {
    fetch('netCrypto.xlsx') // Путь к вашему файлу Excel
    .then(response => response.arrayBuffer())
    .then(buffer => {
        // Парсинг Excel файла
        const workbook = XLSX.read(new Uint8Array(buffer), {type:'array'});

        // Получение первого листа
        const sheet = workbook.Sheets[workbook.SheetNames[0]];

        // Обработка данных
        let totalDeposit = 0;
        let cryptoTotal = 0;
        let roobicTotal = 0;

        const range = XLSX.utils.decode_range(sheet['!ref']);

        for (let rowNum = range.s.r; rowNum <= range.e.r; rowNum++) {
            const depositWithdrawal = sheet['E' + (rowNum + 1)]?.v;
            const amount = sheet['T' + (rowNum + 1)]?.v || 0;
            const cryptoType = sheet['N' + (rowNum + 1)]?.v;
            const roobicType = sheet['P' + (rowNum + 1)]?.v;

            if (cryptoType === 'Crypto' && depositWithdrawal === 'Deposit') {
                cryptoTotal += amount;
            }

            if (roobicType === 'Roobic') {
                roobicTotal += amount;
            }

            if (depositWithdrawal === 'Deposit') {
                totalDeposit += amount;
            } 
        }

        const cryptoPercentage = ((cryptoTotal / totalDeposit) * 100).toFixed(0);
        const roobicPercentage = ((roobicTotal / totalDeposit) * 100).toFixed(0);

        console.log(totalDeposit);
        console.log(cryptoTotal);
        console.log(roobicTotal);

        // Отображение результатов на странице
        document.getElementById('netDeposit').innerText = '$' + totalDeposit.toFixed(0);
        document.getElementById('cryptoPercentage').innerText = '₿' + cryptoPercentage + '%';
        document.getElementById('roobicPercentage').innerText = 'R' + roobicPercentage + '%';
    })
    .catch(error => console.error(error));
};
