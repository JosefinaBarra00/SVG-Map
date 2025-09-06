// Constantes para uso en todo el archivo
const DRIVE_FOLDER_ID = "1Q_iWdMvciaEtKsQdxfg2azQBN7UNr3QP";
const FILE_MAPAS_SKU = "mapaSKU.json";
const FILE_VENCIMIENTOS = "datosVencimientos.json";
const FILE_LOGO_CCU_ID = "1DaPAOxXwksLIv-8-re2GX7jyvUct8opu";

function guardarJSONEnDrive(data, fileName, folderId) {
  try {
    // Convertir a cadena JSON
    const jsonContent = JSON.stringify(data);

    // Obtener la carpeta
    const folder = DriveApp.getFolderById(folderId);

    // Eliminar archivo existente con el mismo nombre
    const existingFiles = folder.getFilesByName(fileName);
    while (existingFiles.hasNext()) {
      existingFiles.next().setTrashed(true);
    }

    // Eliminar también cualquier otro archivo de inventario anterior
    const allFiles = folder.getFiles();
    while (allFiles.hasNext()) {
      const file = allFiles.next();
      if (file.getName().startsWith("Inventario_")) {
        file.setTrashed(true);
      }
    }

    // Crear y guardar el nuevo archivo
    const file = folder.createFile(fileName, jsonContent, MimeType.PLAIN_TEXT);

    // Configurar como público
    file.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);

    Logger.log(`${fileName} guardado exitosamente en Drive`);
    return file.getId();
  } catch (error) {
    Logger.log(`Error guardando ${fileName}: ${error.toString()}`);
    return null;
  }
}
