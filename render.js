const {PythonShell} =  require('python-shell');
const {app, BrowserWindow} = require('electron');
const path = require('path');
//filesystem module
const fs = require("fs");
//Dialogs module
const {dialog} = require('electron').remote;

const { promisify } = require('util');
const sleep = promisify(setTimeout);


const readBtn = document.getElementById('readBtn');
const scrapeBtn = document.getElementById('scrapeBtn');


async function openFile(){
	
	alert("Selecciona la carpeta donde se encuentren los XML")

	//Opens file dialog, looking for xml folder
	openOptions = {
		properties: ['openDirectory']
	}
	const files = await dialog.showOpenDialog(openOptions);

	// If no files.
	if(!files){
		console.log("No files");
		return;
	} 

	console.log(files.filePaths)

	file = files.filePaths[0]

	//const fileContent = fs.readFileSync(file).toString();  //
	//console.log(fileContent)
	console.log(file)
	//options for python
	let options = {
		args:[file]
	}

//	fileWritePath = file.substring(0,2) + "\\" +file.substring(3,file.length-10) + "_elec.xml";
//	fs.writeFile(fileWritePath, fileContent, function(err){
//		if(err) return console.log(err);
//	});	
	python_path = path.join(__dirname, 'parse_xml.py');

	PythonShell.run(python_path,options,  function  (err, results)  {
	 if  (err)  throw err;
	 console.log('python code finished.');
	 console.log('python_results', results);
	 alert("Los valores encontrados estan en el archivo Recibos.xlsx en la misma carpeta")
	});

	
	

}


async function scrapeCFE(){
	
	alert("Selecciona la carpeta donde quieres descargar los XML")

	//Opens file dialog, looking for xml folder
	openOptions = {
		properties: ['openDirectory']
	}
	const files = await dialog.showOpenDialog(openOptions);

	// If no files.
	if(!files){
		console.log("No files");
		return;
	} 

	console.log(files.filePaths)

	file = files.filePaths[0]

	//const fileContent = fs.readFileSync(file).toString();  //
	//console.log(fileContent)
	console.log(file)
	//options for python
	let options = {
		args:[file]
	}

//	fileWritePath = file.substring(0,2) + "\\" +file.substring(3,file.length-10) + "_elec.xml";
//	fs.writeFile(fileWritePath, fileContent, function(err){
//		if(err) return console.log(err);
//	});	
	python_path_cfe = path.join(__dirname, 'get_data_from_cfe.py');

	
	PythonShell.run(python_path_cfe, options,  function  (err, results)  {
	 if  (err)  throw err;
	 console.log('scrape code finished.');
	 console.log('python_results', results);
	});

	

}

async function resetDependencies(_query, sucessCallback,errorCallback){
	python_path_dependencies = path.join(__dirname, 'instalar_librerias_necesarias.py');
	
	PythonShell.run(python_path_dependencies,{}, function  (err, results)  {
	 	if  (err)  throw err;
	 	console.log('dependency code finished.');
	 	console.log('python_results', results);
		alert("Las dependencias se han reiniciado")
	});

}


readBtn.onclick = e => {
	fileContent = openFile();
};

scrapeBtn.onclick = e => {
	fileContent = scrapeCFE();
};


resetBtn.onclick = () => {
	fileContent = resetDependencies();
};
