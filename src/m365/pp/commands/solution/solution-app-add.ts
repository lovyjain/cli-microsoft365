import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import { validation } from '../../../../utils/validation.js';
import PowerAppsCommand from '../../../base/PowerAppsCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  environmentName: string;
  solutionId?: string;
  solutionName?: string;
  appId?: string;
  appName?: string;
  asAdmin?: boolean;
}

class PpSolutionAppAddCommand extends PowerAppsCommand {
  public get name(): string {
    return commands.SOLUTION_APP_ADD;
  }

  public get description(): string {
    return 'Adds an existing Power Apps canvas app to a Power Platform solution';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        solutionId: typeof args.options.solutionId !== 'undefined',
        solutionName: typeof args.options.solutionName !== 'undefined',
        appId: typeof args.options.appId !== 'undefined',
        appName: typeof args.options.appName !== 'undefined',
        asAdmin: !!args.options.asAdmin
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-e, --environmentName <environmentName>'
      },
      {
        option: '--solutionId [solutionId]'
      },
      {
        option: '--solutionName [solutionName]'
      },
      {
        option: '--appId [appId]'
      },
      {
        option: '--appName [appName]'
      },
      {
        option: '--asAdmin'
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['solutionId', 'solutionName'] },
      { options: ['appId', 'appName'] }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.solutionId && !validation.isValidGuid(args.options.solutionId)) {
          return `${args.options.solutionId} is not a valid GUID`;
        }

        if (args.options.appId && !validation.isValidGuid(args.options.appId)) {
          return `${args.options.appId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const solutionId = await this.getSolutionId(args, logger);
      const appId = await this.getAppId(args, logger);

      if (this.verbose) {
        await logger.logToStderr(`Adding app '${args.options.appId || args.options.appName}' to solution '${args.options.solutionId || args.options.solutionName}'...`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/providers/Microsoft.PowerApps/environments/${formatting.encodeQueryParameter(args.options.environmentName)}/apps/${appId}/makeSolutionAware?api-version=2021-02-01`,
        headers: {
          accept: 'application/json'
        },
        responseType: 'json',
        data: {
          solutionId: solutionId
        }
      };

      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getSolutionId(args: CommandArgs, logger: Logger): Promise<string> {
    if (args.options.solutionId) {
      return args.options.solutionId;
    }

    if (this.verbose) {
      await logger.logToStderr(`Retrieving solution '${args.options.solutionName}'...`);
    }

    const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environmentName, args.options.asAdmin);
    const solution = await powerPlatform.getSolutionByName(dynamicsApiUrl, args.options.solutionName!);
    return solution.solutionid;
  }

  private async getAppId(args: CommandArgs, logger: Logger): Promise<string> {
    if (args.options.appId) {
      return args.options.appId;
    }

    if (this.verbose) {
      await logger.logToStderr(`Retrieving app '${args.options.appName}'...`);
    }

    const url = `${this.resource}/providers/Microsoft.PowerApps/scopes/admin/environments/${formatting.encodeQueryParameter(args.options.environmentName)}/apps?api-version=2017-08-01`;
    const apps = await odata.getAllItems<{ name: string; properties: { displayName: string } }>(url);
    const matching = apps.filter(a => a.properties.displayName.toLowerCase() === args.options.appName!.toLowerCase());

    if (matching.length === 0) {
      throw Error(`The specified app '${args.options.appName}' does not exist.`);
    }

    if (matching.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('name', matching);
      const app = await cli.handleMultipleResultsFound<{ name: string; properties: { displayName: string } }>(`Multiple apps with name '${args.options.appName}' found.`, resultAsKeyValuePair);
      return app.name;
    }

    return matching[0].name;
  }
}

export default new PpSolutionAppAddCommand();
