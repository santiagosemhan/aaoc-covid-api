import Excel from 'exceljs';

const exportMedicalRecords = async (user: Parse.User): Promise<Parse.Object> => {
  const medicalRecords = await new Parse.Query('MedicalRecord').find({
    sessionToken: user.getSessionToken(),
  });
  // if (medicalRecords.length) {
  // }
  console.log(medicalRecords);

  const header2 = [
    'ID',
    'Registro Hospitalario',
    'Iniciales',
    'Sexo',
    'Ámbito de atención',
    'Fecha de Nacimiento',
    'País de Nacimiento',
    'Zona de residencia',
    'Estado Civil',
    'Raza',
    'Ocupación Actual',
    'Nivel de Educación',
    'País de Residencia',
    'Peso',
    'Talla',
    'IMC',
    'Tipo Tumoral ',
    'Histología Tumoral',
    'Estadio tumoral',
    'Sitios de metástasis (todos)',
    'Tipo de tratamiento',
    'Describir terapia Oncológica Actual',
    'Objetivo del tratamiento',
    'Si está en tratamiento sistémico, número total de meses en todos los tratamientos hasta el diagnostico de COVID19',
    'Si está en tratamiento sistémico, número total de meses en el último tratamiento hasta el diagnostico de COVID19',
    'Fecha del ultimo tratamiento sistémico',
    'PS al inicio del ultimo tratamiento sistémico',
    'Numero de tratamientos sistémicos previos (solo en enfermedad metastásica)',
    'Tipo de tratamiento al momento del diagnostico de COVID',
    'Si esta con Inmunoterapia , Tipo de Checkpoint inhibitor',
    'Fecha de la primera dosis de Inmunoterapia ',
    'Fecha de la ultima dosis de Inmunoterapia',
    'Efectos adversos relacionados con la Inmunoterapia ',
    'Tipo de Efecto Adverso',
    'Manejo de los Efectos Adversos',
    'Tratamiento del Efecto adverso',
    'Si otros , especifique',
    'Fecha de la cirugía ( si aplica)',
    'Recibió radiación?',
    'Habito Tabáquico',
    'Si fuma, cantidad de paquetes al año',
    'Consumo de alcohol',
    'Hipertensión arterial',
    'Hipercolesterolemia',
    'Obesidad (BMI >=30)',
    'Enfermedad autoinmune',
    'Enfermedad renal crónica',
    'EPOC',
    'Diabetes',
    'Asma',
    'Enfermedad cardiovascular distinta a hipertensión',
    'HIV',
    'Otras comorbilidadesrelevantes',
    'Vacunación previa',
    'Vacuna Influenza 2019',
    'Vacuna Influenza 2020',
    'Medicación habitual',
    'Historia de Hepatitis',
    'Actividad física',
    'Si realiza , describir frecuencia semanal',
    '¿Tuvo una infección respiratoria viral reciente (dentro de 3 meses)?',
    'Uso de anti-inflamatorios?',
    'Tipo de antiinflamatorio que usa',
    'Uso de antibióticos?',
    'Si utiliza , describa cuál antibiótico',
    'Cómo fue diagnosticado el paciente con COVID-19?',
    'Fecha de diagnosis de COVID19/otros virus',
    'Método diagnostico',
    'Si es un método cuantitativo, mencionar valores',
    'Otros métodos ( describir)',
    'Tuvo contacto con una persona sintomática y sospechosa?',
    'Tuvo contacto con una persona con diagnostico conocido de COVID-19 ?',
    'Síntomas al diagnostico',
    'Fecha del inicio de los síntomas ',
    'Tiene otra infección viral (pts. no con COVID-19)?',
    'Tiene otra infección viral junto con COVID-19?',
    'Cual?',
    'Síntomas al diagnostico',
    'Fecha del inicio de los síntomas ',
    'Fecha de diagnóstico de infección por el virus no COVID19',
    'Muestra biológica para diagnóstico',
    'Infección del tracto respiratorio superior en el momento del diagnóstico',
    'Infección del tracto respiratorio bajo en el momento del diagnóstico',
    'El tracto superior evolucionó a infección del tracto respiratorio bajo?',
    'En caso afirmativo, tiempo de infección del tracto respiratorio bajo (días)',
    'Tipo de infección',
    'Hallazgos en Rx de tórax',
    'Hallazgos en TC de Tórax',
    'Terapia con corticoides',
    'Hemoglobina (Hbg)',
    'Volumen corpuscular medio (MCV)',
    'Recuento de plaquetas',
    'Glóbulos blancos al diagnostico (+/-2 días) células/mm3',
    'Neutrófilos al diagnostico (+/-2 días) células/mm3',
    'Linfocitos al diagnostico (+/-2 días) células/mm3',
    'Nivel de Creatinina al diagnostico (+/-2 días) células/mm3',
    'Lactato deshidrogenasa (LDH)',
    'Dímero-D',
    'El paciente fue internado?',
    'Fecha de la internación',
    'Score de severidad? NEWS, Sofa? Precisa? ',
    'Duración (en días) de la internación',
    'Admisión a la UCI al inicio',
    'Admisión en la UCI más tarde durante la enfermedad',
    'Duración ( en días ) de la admisión a UCI',
    'Requirió asistencia respiratoria mecánica',
    'Nivel de saturación de O2 al diagnostico',
    'Temperatura (°C)',
    'Se usó suplemento de oxígeno durante la enfermedad?',
    'Qué se utilizó para el suplemento de O2?',
    'El paciente fue dado de alta con suplementos de O2?',
    'Se usó la terapia antiviral durante la enfermedad?',
    'Se usó oseltamivir para el tratamiento?',
    'Duración de la terapia con oseltamivir',
    'Se uso hidroxicloroquina para el tratamiento?',
    'Duración de la terapia con hidroxicloroquina',
    'Se usó azitromicina para el tratamiento?',
    'Duración de la terapia con azitromicina',
    'Fueron utilizados otros antibióticos para el tratamiento?',
    'Nombre del antibiótico',
    'Duración de la terapia con antibiótico',
    'Se uso tocilizumab para el tratamiento?',
    'Fecha de la administración de tocilizumab',
    'Coinfección respiratoria dentro de las 2 semanas previas a la infección viral?',
    'Tipo de coinfección respiratoria',
    'Lista de microorganismos específicos.',
    'Fecha del ultimo control',
    'El paciente murió?',
    'Fecha de muerte',
    'Cual fue la causa de muerte?',
    'Estado a los 30 días del diagnóstico de infección viral (COVID19 y NONCOVID19)',
    'Estado a los 3 meses',
    'Estado a los 6 meses',
  ];
  // generate workbook
  const workbook = new Excel.Workbook();
  const worksheet = workbook.addWorksheet('Datos registros - covid');
  worksheet.properties.defaultColWidth = 30;

  worksheet.mergeCells('A1:P1');
  worksheet.mergeCells('Q1:AM1');
  worksheet.mergeCells('AN1:BM1');
  worksheet.mergeCells('BN1:EC1');
  worksheet.getCell('A1').value = 'DATOS FILIATORIOS';
  worksheet.getCell('Q1').value = 'CARACTERISTICAS TUMORALES';
  worksheet.getCell('AN1').value = 'INFORMACION DE SALUD';
  worksheet.getCell('BN1').value = 'COVID-19 / OTROS VIRUS';

  const rows = [header2, [0, 1]];
  worksheet.addRows(rows);

  worksheet.getCell('A1').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFFFFF00' },
    bgColor: { argb: 'FF0000FF' },
  };
  worksheet.getCell('Q1').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: '00B050' },
    bgColor: { argb: '00B050' },
  };
  worksheet.getCell('AN1').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'BDD7EE' },
    bgColor: { argb: 'BDD7EE' },
  };
  worksheet.getCell('BN1').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF0000' },
    bgColor: { argb: 'FF0000' },
  };

  worksheet.getColumn(1).width = 12;

  worksheet.getRow(1).alignment = { horizontal: 'center' };
  worksheet.getRow(1).height = 15;
  worksheet.getRow(1).font = { bold: true };

  worksheet.getRow(2).font = { bold: true };
  worksheet.getRow(2).height = 35;
  worksheet.getRow(2).alignment = { wrapText: true };

  const buffer = await workbook.xlsx.writeBuffer();

  // Create a Parse File and attach to Report Object.
  const date = new Date().toLocaleDateString('es-AR');
  const strDate = date.replace(/\//g, '-');
  const fileExport = new Parse.Object('Report');
  const file = new Parse.File(`reporte-${strDate}.xlsx`, Array.from(<Buffer>buffer));
  fileExport.set('file', file);
  fileExport.set('type', 'xlsx');
  await fileExport.save(null, { useMasterKey: true });
  return fileExport;
};

export default { exportMedicalRecords };
