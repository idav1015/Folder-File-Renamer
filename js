/* ------- CHANGE VARIABLES ------- */

const parentFolderId = "1nl-kHSYOb6LnqG50K5E-xOfI2Ays7EHM"; //Parent folder ID (starting folder)
const currentFolderName = "NEW CAFM TY"; //CAFM 001 folder name (to change); 
const newFolderName = "CAFM TY" //CAMF 001 new folder name

let counter1 = 0;
let counter2 = 0;
let foldersToRename = [];
let foldersRenamed = []


class Folder {
  constructor(id, name, link) {
    this.id = id;
    this.name = name;
    this.link = link;
    console.log(`ID: ${id}, name: ${name}, link: ${link}`)
  }
}

function main() {
  const foldersToRename = getFolders(parentFolderId, currentFolderName); 
  renameFolders(foldersToRename, newFolderName)
  createSheet(foldersToRename, foldersRenamed);

}


function getFolders(startFolder, searchFolder){
  const folder = DriveApp.getFolderById(startFolder);
  const subFolders = folder.getFolders();
  const folderName = folder.getName().toUpperCase();
  

  if(folderName.includes(currentFolderName.toUpperCase())){
    counter1 += 1;
    const link =  folder.getUrl();
    const folderObject = new Folder(folder.getId(), folder.getName(), link);
    foldersToRename.push(folderObject);
    console.log("folders to rename: " + counter1);
  }

  while(subFolders.hasNext()){
    const subfolder = subFolders.next().getId();
    getFolders(subfolder, searchFolder);
  }
  return foldersToRename;
}

function renameFolders(folders, newName) {
  folders.forEach(folder => {
    const oldName = folder.name
    const folderObject = DriveApp.getFolderById(folder.id);
    const newFolder = folder.name.replace(folder.name, newName);
    folderObject.setName(newFolder); // Renaming the actual folder
    const newLink = folderObject.getUrl();
    const newFolderObject = new Folder(folder.id, newFolder, newLink);
    foldersRenamed.push(newFolderObject);

    counter2 += 1 
    console.log(counter2)

    console.log(`Renamed folder: ${oldName} to ${newFolder}`);
  })

  return foldersRenamed  

}


function createSheet(currentFolders, newFolders) {
  const spreadsheet = SpreadsheetApp.create("Folders List");
  const sheet1 = spreadsheet.getActiveSheet();
  sheet1.setName("Current Folders");
  sheet1.appendRow(["ID", "Name", "Link"]);

  currentFolders.forEach(folder => {
    sheet1.appendRow([folder.id, folder.name, folder.link]);
  });

  const sheet2 = spreadsheet.insertSheet('Renamed Folders');
  sheet2.appendRow(["ID", "Name", "Link"]);

  
  newFolders.forEach(folder => {
    sheet2.appendRow([folder.id, folder.name, folder.link]);
  });
  console.log("DONE")
}


