const data = require('./index.json');
const excel = require('node-excel-export');
const fs = require('fs');

(() => {
    const months = [
        'Enero',
        'Febrero',
        'Marzo',
        'Abril',
        'Mayo',
        'Junio',
        'Julio',
        'Agosto',
        'Septiembre',
        'Octubre',
        'Noviembre',
        'Diciembre'
    ];

    // You can define styles as json object
    const styles = {
        headerDark: {
            // fill: {
            //     fgColor: {
            //         rgb: 'FF000000'
            //     }
            // },
            font: {
                // color: {
                //     rgb: 'FFFFFFFF'
                // },
                sz: 13,
                bold: true,
                // underline: true
            }
        },
        cellPink: {
            fill: {
                fgColor: {
                    rgb: 'FFFFCCFF'
                }
            }
        },
        cellGreen: {
            fill: {
                fgColor: {
                    rgb: 'FF00FF00'
                }
            }
        }
    };

    //Array of objects representing heading rows (very top)
    const heading = [];

    //Here you specify the export structure
    const specification = {
        tipo: {
            displayName: 'tipo',
            headerStyle: styles.headerDark,
            width: '15'
        },
        year: {
            displayName: 'year',
            headerStyle: styles.headerDark,
            width: '15'
        },
        month: {
            displayName: 'month',
            headerStyle: styles.headerDark,
            width: '15'
        },
        provincia: {
            displayName: 'provincia',
            headerStyle: styles.headerDark,
            width: '15'
        },
        /* letra_provincia_transferencia: {
            displayName: 'letra_provincia_transferencia',
            headerStyle: styles.headerDark,
            width: '15'
        }, */
        pantentamientos: {
            displayName: 'pantentamientos',
            headerStyle: styles.headerDark,
            width: '15'
        },
        /* provincia_id: {
            displayName: 'provincia_id',
            headerStyle: styles.headerDark,
            width: '15'
        } */
    }

    const dataset = [];

    let year = '', name = '', month = 0;

    for (let i = 1; i <= 5; i++) {
        const sheetData = data[`2.${i}`];
        for (let j = 0; j < sheetData.length; j++) {
            name = sheetData[j]?.['name'] ? sheetData[j]['name'] : name;
            // Filtering month for any astric or any other text other than number
            let tempYear;
            if (sheetData[j]?.['year']) {
                tempYear = tempYear = String(sheetData[j]?.['year']).replace('*', '');
            }
            year = tempYear || year;
            console.log(`Processing Year: ${year}`);

            // Filtering month for any astric
            let tempMonth;
            if (sheetData[j]?.['month']) {
                tempMonth = String(sheetData[j]?.['month']).replace('*', '');
            }
            month = tempMonth ? months.indexOf(tempMonth) : months.indexOf(month);
            console.log(`Processing Year-Month: ${year}-${month + 1}`);

            if (sheetData[j]) {
                for (const [key, value] of Object.entries(sheetData[j])) {
                    // console.log(`key: ${key}`);

                    if (key === 'year') console.log("🚀 ~ file: index.js:113 ~ key:", key, sheetData[j])
                    if (!['year', 'name', 'month', 'total'].includes(key)) {
                        const tempObject = {
                            'tipo': name,
                            'year': year,
                            'month': month + 1,
                            'provincia': key,
                            // 'letra_provincia_transferencia': '',
                            'pantentamientos': value,
                            // 'provincia_id': ''
                        };
                        dataset.push(tempObject);
                    }
                }
            }
        }
    }
    const merges = []

    const report = excel.buildExport(
        [ // <- Notice that this is an array. Pass multiple sheets to create multi sheet report
            {
                name: 'Sheet-1', // <- Specify sheet name (optional)
                heading: heading, // <- Raw heading array (optional)
                merges: merges, // <- Merge cell ranges
                specification: specification, // <- Report specification
                data: dataset // <-- Report data
            }
        ]
    );
    fs.writeFileSync('patentamientos-new.xlsx', report);
})();
