import logo from './logo.svg';
import './App.css';
import writeXlsxFile from 'write-excel-file'

const columns = [
  { width: 4.83 },
  { width: 10.83 },
  { width: 7.33 }, // in characters
  { width: 4.83 },
  { width: 4.83 },
  { width: 4.83 },
  { width: 6.50 },
  { width: 11.17 },
  { width: 6.83 },
  { width: 8.83 },
  { width: 9.50 },
  { width: 9.50 },
  { width: 9.50 },
  { width: 11.50 }
]

const TITLE_ROW = [
  {
    value: 'CHEQUEO DE MASTITIS',
    fontWeight: 'bold',
    align: 'center',
    alignVertical: 'center',
    height: 36,
    fontSize: 35,
    color: '#FFFFFF',
    backgroundColor: '#00B3BC',
    span: 13,
    borderColor: '#000000'
  },
  null,
  null,
  null,
  null,
  null,
  null,
  null,
  null,
  null,
  null,
  null,
  null,
  {
    value: 6,
    fontWeight: 'bold',
    align: 'center',
    alignVertical: 'center',
    height: 36,
    fontSize: 35,
    color: '#FFFFFF',
    backgroundColor: '#00B3BC',
    borderColor: '#000000'
  }
]

const FARM_INFO_ROW_1 = [
  {
    value: 'PROPIETARIO',
    align: 'center',
    alignVertical: 'center',
    fontWeight: 'bold',
    color: '#FFFFFF',
    backgroundColor: '#00B3BC',
    fontSize: 11,
    span: 2,
    borderColor: '#000000'
  },
  null,
  {
    value: '',
    align: 'center',
    alignVertical: 'center',
    fontWeight: 'bold',
    fontSize: 11,
    span: 8,
    borderColor: '#000000'
  },
  null,
  null,
  null,
  null,
  null,
  null,
  null,
  {
    value: '$/Lt:',
    align: 'center',
    alignVertical: 'center',
    fontWeight: 'bold',
    color: '#FFFFFF',
    fontSize: 11,
    backgroundColor: '#00B3BC',
    borderColor: '#000000'
  },
  {
    value: '',
    align: 'center',
    alignVertical: 'center',
    fontSize: 11,
    fontWeight: 'bold',
    borderColor: '#000000'
  },
  {
    value: 'Fecha:',
    align: 'center',
    alignVertical: 'center',
    fontWeight: 'bold',
    color: '#FFFFFF',
    fontSize: 11,
    backgroundColor: '#00B3BC',
    borderColor: '#000000'
  },
  {
    value: '',
    align: 'center',
    alignVertical: 'center',
    fontSize: 11,
    fontWeight: 'bold',
    borderColor: '#000000'
  }
]

const DATA_ROW_1 = [
  // "Name"
  {
    type: Number,
    value: 2,
    align: 'center',
    alignVertical: 'center',
    color: '#FFFFFF',
    backgroundColor: '#00B3BC'
  },

  {
    type: Number,
    align: 'center',
    alignVertical: 'center',
    value: 953
  },

  {
    type: String,
    align: 'center',
    alignVertical: 'center',
    value: ''
  },
  {
    type: String,
    align: 'center',
    alignVertical: 'center',
    value: ''
  },
  {
    type: String,
    align: 'center',
    alignVertical: 'center',
    value: ''
  },
  {
    type: String,
    align: 'center',
    alignVertical: 'center',
    value: ''
  },

  {
    type: Number,
    align: 'center',
    alignVertical: 'center',
    format: '#,#0.0',
    value: 0.0
  },
  {
    type: String,
    align: 'center',
    alignVertical: 'center',
    value: ''
  },
  {
    type: Number,
    align: 'center',
    alignVertical: 'center',
    color: '#9C6500',
    backgroundColor: '#FFEB9C',
    value: 4
  },
  {
    type: Number,
    align: 'center',
    alignVertical: 'center',
    color: '#336633',
    backgroundColor: '#CCFFCC',
    value: 24
  },
  {
    type: String,
    align: 'center',
    alignVertical: 'center',
    value: ''
  },
  {
    type: String,
    align: 'center',
    alignVertical: 'center',
    value: ''
  },
  {
    //type: Number,
    align: 'center',
    alignVertical: 'center',
    //format: '#,#0.0',
    color: '#336633',
    backgroundColor: '#CCFFCC',
    value: '=4*3'
  },
  {
    type: Number,
    align: 'center',
    alignVertical: 'center',
    color: '#336633',
    backgroundColor: '#CCFFCC',
    value: 1
  }
]

const FOOTER_ROW_1 = [
  {
    value: 1116,
    fontWeight: 'bold',
    align: 'center',
    alignVertical: 'center',
    //height: 36,
    fontSize: 35,
    span: 2,
    rowSpan: 4,
    borderColor: '#000000'
  },
  null
];

const FOOTER_ROW_2 = [
  null,
  null
];

const FOOTER_ROW_3 = [
  null,
  null
];

const FOOTER_ROW_4 = [
  null,
  null
];

const data = [
  TITLE_ROW,
  FARM_INFO_ROW_1,
  DATA_ROW_1,
  FOOTER_ROW_1,
  FOOTER_ROW_2,
  FOOTER_ROW_3,
  FOOTER_ROW_4
];

function generateExcel() {
  return writeXlsxFile(data, {
    columns, // (optional) column widths, etc.
    fileName: 'file.xlsx',
    stickyRowsCount: 1,
    stickyColumnsCount: 14
  }).catch((err) => {
    console.log('==> Se murio algo: ', err);
  });
}

function App() {
  return (
    <div className="App">
      <header className="App-header">
        <img src={logo} className="App-logo" alt="logo" />
        <p>
          Edit <code>src/App.js</code> and save to reload.
        </p>
        <a
          className="App-link"
          href="https://reactjs.org"
          target="_blank"
          rel="noopener noreferrer"
        >
          Learn React
        </a>
        <button onClick={generateExcel}>Excel</button>
      </header>
    </div>
  );
}

export default App;
