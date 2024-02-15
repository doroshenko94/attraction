window.onload = function() {
    fetch('netCrypto.xlsx') // Путь к вашему файлу Excel
    .then(response => response.arrayBuffer())
    .then(buffer => {
        // Парсинг Excel файла
        const workbook = XLSX.read(new Uint8Array(buffer), {type:'array'});

        // Получение первого листа
        const sheet = workbook.Sheets[workbook.SheetNames[0]];

        // Обработка данных
        let grossDeposit = 0;
        let grossWithdrawal = 0; // Добавляем переменную для общей суммы выводов
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
                grossDeposit += amount;
            } else if (depositWithdrawal === 'Withdrawal') { // Добавляем условие для выводов
                grossWithdrawal += amount;
            }
        }
        
        const netDeposit = grossDeposit - grossWithdrawal; // Вычисляем чистый депозит

        const cryptoPercentage = ((cryptoTotal / grossDeposit) * 100).toFixed(0);
        const roobicPercentage = ((roobicTotal / grossDeposit) * 100).toFixed(0);

        console.log(grossDeposit);
        console.log(cryptoTotal);
        console.log(roobicTotal);

        // Отображение результатов на странице
        document.getElementById('netDeposit').innerText = '$' + netDeposit.toFixed(0); // Отображаем чистый депозит
        document.getElementById('cryptoPercentage').innerText = '₿' + cryptoPercentage + '%';
        document.getElementById('roobicPercentage').innerText = 'R' + roobicPercentage + '%';
    })
    .catch(error => console.error(error));

    // Загрузка данных и обновление страницы при загрузке и каждый час
    loadDataAndRefresh(); // Выполняем сразу после загрузки страницы
    setInterval(loadDataAndRefresh, 3600000); // 3600000 миллисекунд = 1 час
};
