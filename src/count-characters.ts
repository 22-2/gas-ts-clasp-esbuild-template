import { NOVEL_FOLDER_ID } from './constants';
import { parseMd } from './utils';

function getFilesInFolder(folderId: string) {
  const folder = DriveApp.getFolderById(folderId);
  const filesIt = folder.getFiles();
  const ret = [];
  while (filesIt.hasNext()) {
    ret.push(filesIt.next());
  }
  return ret;
}

export function getEpisodeFiles() {
  return [...getFilesInFolder(NOVEL_FOLDER_ID)];
}

export function requestCharCount() {
  let sum = 0;
  for (const file of getEpisodeFiles()) {
    const content = file.getBlob().getDataAsString('utf-8');
    const parsed = parseMd(content);
    sum += parsed.content.replace(/[\s\n]+/g, '').length;
  }
  return sum;
}
