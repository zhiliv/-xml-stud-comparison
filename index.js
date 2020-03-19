'use strict';

//модуль для конвертирования эксель в json
const excelToJson = require('convert-excel-to-json'),
  //преобразование данных в excel
  json2xls = require('json2xls'),
  //модуль для работы с промисами
  q = require('q'),
  //модуль для работы с асинхронными модулями
  async = require('async'),
  //модуль для работы с файловой системой
  fs = require('fs'),
  //модуль для работы с данными
  _ = require('underscore');

/************************************************
 * Получение зарегистрирвоанных преподавателей  *
 * @function getRegPrep                         *
 ************************************************/
const getRegPrep = async () => {
  let defer = q.defer();
  defer.resolve(
    excelToJson({
      sourceFile: 'prep.xlsx'
    })
  );
  return defer.promise;
};

/**********************************************************
 * Получение данных сотрудников старооскольского филиала  *
 * @function getST                                        *
 **********************************************************/
const getST = async () => {
  let defer = q.defer();
  defer.resolve(
    excelToJson({
      sourceFile: 'ST.xlsx'
    })
  );
  return defer.promise;
};

/**********************************************************
 * Получение данных сотрудников старооскольского филиала  *
 * @function getGF                                        *
 **********************************************************/
const getGF = async () => {
  let defer = q.defer();
  defer.resolve(
    excelToJson({
      sourceFile: 'GF.xls'
    })
  );
  return defer.promise;
};

getRegPrep().then(async regPrep => {
  //начало с  regPrep.TDSheet[3]
  //хранение результата
  let arr = [];
  //список старого оскола
  let stList;
  //список губкина
  let gfList;
  await getST().then(st => {
    stList = st['Sheet 1'];
  });
  await getGF().then(gf => {
    //начало с gf.TDSheet[2]
    gfList = gf.TDSheet;
  });
  //обход всех значений старооскольского филиала
  console.log(stList[4])
  await async.eachOfSeries(stList, async (row, ind) => {
    if (ind > 0) {
      let famalyST = String(row.A);
      let nameST = String(row.B);
      let otchST = String(row.C);
      let obj = { H: famalyST, I: nameST, J: otchST };
      let chk = _.where(regPrep['Лист2'], obj);
      if (chk.length == 0) {
        if (String(row.F) == 'Основное место работы' && String(row.F == 'Внешнее совместительство')) {
          let FIO = String(row.B).split(' ');
          let obj = {
            'Фамилия': famalyST,
            'Имя': nameST,
            'Отчество': otchST,
            'Должность': row.D,
            'Подразделение': row.E,
            'Вид занятости': row.F,
            'Ставка': row.G,
            'Таб №': String(row.H),
            'Паспортные данные': String(row.I)
          };
          let chkAr = _.where(arr, obj)
          if(chkAr.length == 0){
            arr.push(obj);
          }
          
        }
      }
    }
  });

  await async.eachOfSeries(gfList, async (row, ind) => {
    if (ind > 0) {
      let famalyST = String(row.B);
      let nameST = String(row.C);
      let otchST = String(row.D);
      let obj = { H: famalyST, I: nameST, J: otchST };
      let chk = _.where(regPrep['Лист2'], obj);
      if (chk.length == 0) {
        if (String(row.F) == 'Основное место работы' && String(row.F) == 'Внешнее совместительство') {
          let FIO = String(row.B).split(' ');
          let obj = {
            Фамилия: FIO[0],
            Имя: FIO[1],
            Отчество: FIO[2],
            Должность: row.C,
            Подразделение: row.D,
            'Вид занятости': row.E,
            Ставка: row.F,
            'Таб №': String(row.G)
          };
          arr.push(obj);
        }
      }
    }
  });
  var xls = json2xls(arr);
  fs.writeFileSync('data.xlsx', xls, 'binary');
  console.log(arr.length)
});
