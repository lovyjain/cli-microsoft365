import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import { validation } from '../../../../utils/validation.js';
import PowerAutomateCommand from '../../../base/PowerAutomateCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  environmentName: string;
  solutionId?: string;
  solutionName?: string;
  flowId?: string;
  flowName?: string;
  asAdmin?: boolean;
}

class PpSolutionFlowAddCommand extends PowerAutomateCommand {
  public get name(): string {
    return commands.SOLUTION_FLOW_ADD;
  }

  public get description(): string {
    return 'Adds an existing Power Automate cloud flow to a Power Platform solution';
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
        flowId: typeof args.options.flowId !== 'undefined',
        flowName: typeof args.options.flowName !== 'undefined',
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
        option: '--flowId [flowId]'
      },
      {
        option: '--flowName [flowName]'
      },
      {
        option: '--asAdmin'
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['solutionId', 'solutionName'] },
      { options: ['flowId', 'flowName'] }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.solutionId && !validation.isValidGuid(args.options.solutionId)) {
          return `${args.options.solutionId} is not a valid GUID`;
        }

        if (args.options.flowId && !validation.isValidGuid(args.options.flowId)) {
          return `${args.options.flowId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const solutionId = await this.getSolutionId(args, logger);
      const flowId = await this.getFlowId(args, logger);

      if (this.verbose) {
        await logger.logToStderr(`Adding flow '${args.options.flowId || args.options.flowName}' to solution '${args.options.solutionId || args.options.solutionName}'...`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${PowerAutomateCommand.resource}/providers/Microsoft.Flow/environments/${formatting.encodeQueryParameter(args.options.environmentName)}/solutions/${solutionId}/migrateFlows?api-version=2018-10-01`,
        headers: {
          accept: 'application/json'
        },
        responseType: 'json',
        data: {
          flowsToMigrate: [flowId]
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

  private async getFlowId(args: CommandArgs, logger: Logger): Promise<string> {
    if (args.options.flowId) {
      return args.options.flowId;
    }

    if (this.verbose) {
      await logger.logToStderr(`Retrieving flow '${args.options.flowName}'...`);
    }

    const url = `${PowerAutomateCommand.resource}/providers/Microsoft.ProcessSimple/environments/${formatting.encodeQueryParameter(args.options.environmentName)}/flows?api-version=2016-11-01`;
    const flows = await odata.getAllItems<{ name: string; properties: { displayName: string } }>(url);
    const matching = flows.filter(f => f.properties.displayName.toLowerCase() === args.options.flowName!.toLowerCase());

    if (matching.length === 0) {
      throw Error(`The specified flow '${args.options.flowName}' does not exist.`);
    }

    if (matching.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('name', matching);
      const flow = await cli.handleMultipleResultsFound<{ name: string; properties: { displayName: string } }>(`Multiple flows with name '${args.options.flowName}' found.`, resultAsKeyValuePair);
      return flow.name;
    }

    return matching[0].name;
  }
}

export default new PpSolutionFlowAddCommand();
