const ExcelJs = require('exceljs');
const users = require('./users.json');

const createExcel = async (users) => {
    // Crear un archivo excel
    const workbook = new ExcelJs.Workbook();
    // Crear una hoja dentro del archivo excel
    const worksheet = workbook.addWorksheet('Users');

    // Definir las columnas de la hoja del excel
    worksheet.columns = [
        { header: 'Id', key: 'id', width: 10 },
        { header: 'First name', key: 'first_name', width: 20 },
        { header: 'Last name', key: 'last_name', width: 20 },
        { header: 'Email', key: 'email', width: 30 },
        { header: 'Password', key: 'password', width: 20 },
    ];

    // Agregar la data a la hoja del excel
    worksheet.addRows(users);

    // Guardar el excel
    await workbook.xlsx.writeFile('users.xlsx');

    console.log('File created');

    // Cargar un excel
    const newWorkbook = new ExcelJs.Workbook();
    await newWorkbook.xlsx.readFile('users.xlsx');

    // Obtener una hoja del excel cargado
    const newWorksheet = newWorkbook.getWorksheet('Users');
    // Definir las columnas del nuevo archivo
    newWorksheet.columns = [
        { header: 'Id', key: 'id', width: 10 },
        { header: 'First name', key: 'first_name', width: 20 },
        { header: 'Last name', key: 'last_name', width: 20 },
        { header: 'Email', key: 'email', width: 30 },
        { header: 'Password', key: 'password', width: 20 },
    ];

    // Agregar un nuevo row al excel
    newWorksheet.addRow({ id: 11, first_name: 'user 11', last_name: 'last name 11', email: 'usertest@gmail.com', password: '123456' });

    // Guardar el excel
    await newWorkbook.xlsx.writeFile('users2.xlsx');

    console.log('File created');
};

createExcel(users);