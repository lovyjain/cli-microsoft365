import { Drive, DriveItem, Site } from '@microsoft/microsoft-graph-types';
import { AxiosRequestConfig } from 'axios';
import { Logger } from '../../../cli/Logger';
import GlobalOptions from '../../../GlobalOptions';
import request from '../../../request';
import { odata } from '../../../utils/odata';
import { validation } from '../../../utils/validation';
import GraphCommand from '../../base/GraphCommand';
import commands from '../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  folderUrl: string;
  recursive?: boolean;
}

class FileListCommand extends GraphCommand {
  foldersToGetFilesFrom: string[] = [];

  public get name(): string {
    return commands.LIST;
  }

  public get description(): string {
    return 'Retrieves files from the specified folder and site';
  }

  public defaultProperties(): string[] | undefined {
    return ['name', 'lastModifiedByUser'];
  }

  constructor() {
    super();
  
    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }
  
  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        recursive: !!args.options.recursive
      });
    });
  }
  
  #initOptions(): void {
    this.options.unshift(
      { option: '-u, --webUrl <webUrl>' },
      { option: '-f, --folderUrl <folderUrl>' },
      { option: '--recursive' }
    );
  }
  
  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => validation.isValidSharePointUrl(args.options.webUrl)
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let webUrl: string = args.options.webUrl;
    if (!webUrl.endsWith('/')) {
      webUrl += '/';
    }
    const folderUrl: URL = new URL(args.options.folderUrl, webUrl);
    let driveId: string = '';

    try {
      const siteId = await this.getSiteId(args.options.webUrl, logger);
      const drive = await this.getDocumentLibrary(siteId, folderUrl, args.options.folderUrl, logger);
      driveId = drive.id as string;

      const folderId = await this.getStartingFolderId(drive, folderUrl, logger);
      if (this.verbose) {
        logger.logToStderr(`Loading folders to get files from...`);
      }

      // add the starting folder to the list of folders to get files from
      this.foldersToGetFilesFrom.push(folderId);
      await this.loadFoldersToGetFilesFrom(folderId, driveId, args.options.recursive);
      if (this.debug) {
        logger.logToStderr(`Folders to get files from: ${this.foldersToGetFilesFrom.join(', ')}`);
      }

      const files = await this.loadFilesFromFolders(driveId, this.foldersToGetFilesFrom, logger);
      logger.log(files);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getSiteId(webUrl: string, logger: Logger): Promise<string> {
    if (this.verbose) {
      logger.logToStderr(`Getting site id...`);
    }

    const url: URL = new URL(webUrl);
    const requestOptions: AxiosRequestConfig = {
      url: `${this.resource}/v1.0/sites/${encodeURIComponent(url.host)}:${url.pathname}?$select=id`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };
    return request
      .get<Site>(requestOptions)
      .then((site: Site): string => {
        if (this.verbose) {
          logger.logToStderr(`Site id: ${site.id}`);
        }

        return site.id as string;
      });
  }

  private getDocumentLibrary(siteId: string, folderUrl: URL, folderUrlFromUser: string, logger: Logger): Promise<Drive> {
    if (this.verbose) {
      logger.logToStderr(`Getting document library...`);
    }

    const requestOptions: AxiosRequestConfig = {
      url: `${this.resource}/v1.0/sites/${siteId}/drives?$select=webUrl,id`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };
    return request
      .get<{ value: Drive[] }>(requestOptions)
      .then((drives: { value: Drive[] }): Promise<Drive> => {
        const lowerCaseFolderUrl: string = folderUrl.href.toLowerCase();
        const drive: Drive | undefined = drives.value
          .sort((a, b) => (b.webUrl as string).localeCompare(a.webUrl as string))
          .find((d: Drive) => {
            const driveUrl: string = (d.webUrl as string).toLowerCase();
            // ensure that the drive url is a prefix of the folder url
            return lowerCaseFolderUrl.startsWith(driveUrl) &&
              (driveUrl.length === lowerCaseFolderUrl.length ||
                lowerCaseFolderUrl[driveUrl.length] === '/');
          });

        if (!drive) {
          return Promise.reject(`Document library '${folderUrlFromUser}' not found`);
        }

        if (this.verbose) {
          logger.logToStderr(`Document library: ${drive.webUrl}, ${drive.id}`);
        }

        return Promise.resolve(drive);
      });
  }

  private getStartingFolderId(documentLibrary: Drive, folderUrl: URL, logger: Logger): Promise<string> {
    if (this.verbose) {
      logger.logToStderr(`Getting starting folder id...`);
    }

    const documentLibraryRelativeFolderUrl: string = folderUrl.href.replace(new RegExp(documentLibrary.webUrl as string, 'i'), '');
    const requestOptions: AxiosRequestConfig = {
      url: `${this.resource}/v1.0/drives/${documentLibrary.id}/root${documentLibraryRelativeFolderUrl.length > 0 ? `:${documentLibraryRelativeFolderUrl}` : ''}?$select=id`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };
    return request
      .get<DriveItem>(requestOptions)
      .then((folder: DriveItem): string => {
        if (this.verbose) {
          logger.logToStderr(`Starting folder id: ${folder.id}`);
        }

        return folder.id as string;
      });
  }

  private loadFoldersToGetFilesFrom(folderId: string, driveId: string, recursive: boolean | undefined): Promise<void> {
    if (!recursive) {
      return Promise.resolve();
    }

    const requestOptions: AxiosRequestConfig = {
      url: `${this.resource}/v1.0/drives/${driveId}/items('${folderId}')/children?$filter=folder ne null&$select=id`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };
    return request
      .get<{ value: DriveItem[] }>(requestOptions)
      .then((subfolders: { value: DriveItem[] }): Promise<void> => {
        const subfolderIds: string[] = subfolders.value.map((subfolder: DriveItem) => subfolder.id as string);
        this.foldersToGetFilesFrom = this.foldersToGetFilesFrom.concat(subfolderIds);
        return Promise
          .all(subfolderIds.map((subfolderId: string) => this.loadFoldersToGetFilesFrom(subfolderId, driveId, recursive)))
          .then(_ => Promise.resolve());
      });
  }

  private loadFilesFromFolders(driveId: string, folderIds: string[], logger: Logger): Promise<DriveItem[]> {
    if (this.verbose) {
      logger.logToStderr(`Loading files from folders...`);
    }

    let files: DriveItem[] = [];

    return Promise
      .all(folderIds.map((folderId: string): Promise<DriveItem[]> =>
        // get items from folder. Because we can't filter out folders here
        // we need to get all items from the folder and filter them out later
        odata.getAllItems<DriveItem>(`${this.resource}/v1.0/drives/${driveId}/items/${folderId}/children`)))
      .then(res => {
        // flatten data from all promises
        files = files.concat(...res);

        // remove folders from the list of files
        files = files.filter((item: DriveItem) => item.file);
        files.forEach(file => (file as any).lastModifiedByUser = file.lastModifiedBy?.user?.displayName);
        return files;
      });
  }
}

module.exports = new FileListCommand();