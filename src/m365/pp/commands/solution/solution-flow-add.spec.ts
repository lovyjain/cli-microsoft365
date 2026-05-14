import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { odata } from '../../../../utils/odata.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { accessToken } from '../../../../utils/accessToken.js';
import commands from '../../commands.js';
import command from './solution-flow-add.js';

describe(commands.SOLUTION_FLOW_ADD, () => {
  let commandInfo: CommandInfo;

  //#region Mocked Responses
  const validEnvironment = '4be50206-9576-4237-8b17-38d8aadfaa36';
  const validSolutionId = '00000001-0000-0000-0001-00000000009b';
  const validSolutionName = 'MySolution';
  const validFlowId = 'fc58e85b-768d-4e16-88af-7f96496a39be';
  const validFlowName = 'VersionCheck';
  const dynamicsApiUrl = 'https://contoso-dev.api.crm4.dynamics.com';
  const migrateFlowsUrl = `https://api.flow.microsoft.com/providers/Microsoft.Flow/environments/${validEnvironment}/solutions/${validSolutionId}/migrateFlows?api-version=2018-10-01`;

  const solutionResponse = {
    solutionid: validSolutionId,
    uniquename: validSolutionName,
    friendlyname: validSolutionName
  };

  const flowsListResponse = [
    {
      name: validFlowId,
      properties: { displayName: validFlowName }
    }
  ];

  const multipleFlowsResponse = [
    {
      name: validFlowId,
      properties: { displayName: validFlowName }
    },
    {
      name: 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee',
      properties: { displayName: validFlowName }
    }
  ];
  //#endregion

  let log: string[];
  let logger: Logger;
  let loggerLogToStderrSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(accessToken, 'assertAccessTokenType').returns();
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: async (msg: string) => { log.push(msg); },
      logRaw: async (msg: string) => { log.push(msg); },
      logToStderr: async (msg: string) => { log.push(msg); }
    };
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      odata.getAllItems,
      powerPlatform.getDynamicsInstanceApiUrl,
      powerPlatform.getSolutionByName,
      cli.handleMultipleResultsFound
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SOLUTION_FLOW_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct option sets', () => {
    const optionSets = command.optionSets;
    assert.deepStrictEqual(optionSets, [
      { options: ['solutionId', 'solutionName'] },
      { options: ['flowId', 'flowName'] }
    ]);
  });

  it('fails validation if solutionId is not a valid GUID', async () => {
    const actual = await command.validate({
      options: {
        environmentName: validEnvironment,
        solutionId: 'invalid-guid',
        flowId: validFlowId
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if flowId is not a valid GUID', async () => {
    const actual = await command.validate({
      options: {
        environmentName: validEnvironment,
        solutionId: validSolutionId,
        flowId: 'invalid-guid'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when solutionId and flowId are valid GUIDs', async () => {
    const actual = await command.validate({
      options: {
        environmentName: validEnvironment,
        solutionId: validSolutionId,
        flowId: validFlowId
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when solutionName and flowName are provided', async () => {
    const actual = await command.validate({
      options: {
        environmentName: validEnvironment,
        solutionName: validSolutionName,
        flowName: validFlowName
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('adds a cloud flow to a solution using IDs and sends correct payload', async () => {
    let postBody: any;
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === migrateFlowsUrl) {
        postBody = opts.data;
        return;
      }
      throw `Invalid POST request: ${opts.url}`;
    });

    await assert.doesNotReject(command.action(logger, {
      options: {
        environmentName: validEnvironment,
        solutionId: validSolutionId,
        flowId: validFlowId
      }
    }));

    assert.deepStrictEqual(postBody, { flowsToMigrate: [validFlowId] });
  });

  it('adds a cloud flow to a solution by resolving names via Dataverse and ProcessSimple API', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').resolves(dynamicsApiUrl);
    sinon.stub(powerPlatform, 'getSolutionByName').resolves(solutionResponse);
    sinon.stub(odata, 'getAllItems').resolves(flowsListResponse);

    let postBody: any;
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === migrateFlowsUrl) {
        postBody = opts.data;
        return;
      }
      throw `Invalid POST request: ${opts.url}`;
    });

    await assert.doesNotReject(command.action(logger, {
      options: {
        environmentName: validEnvironment,
        solutionName: validSolutionName,
        flowName: validFlowName
      }
    }));

    assert.deepStrictEqual(postBody, { flowsToMigrate: [validFlowId] });
  });

  it('resolves cloud flow name case-insensitively', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').resolves(dynamicsApiUrl);
    sinon.stub(powerPlatform, 'getSolutionByName').resolves(solutionResponse);
    sinon.stub(odata, 'getAllItems').resolves(flowsListResponse);

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === migrateFlowsUrl) {
        return;
      }
      throw `Invalid POST request: ${opts.url}`;
    });

    await assert.doesNotReject(command.action(logger, {
      options: {
        environmentName: validEnvironment,
        solutionName: validSolutionName,
        flowName: validFlowName.toUpperCase()
      }
    }));
  });

  it('adds a cloud flow to a solution as admin', async () => {
    const getDynamicsStub = sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').resolves(dynamicsApiUrl);
    sinon.stub(powerPlatform, 'getSolutionByName').resolves(solutionResponse);
    sinon.stub(odata, 'getAllItems').resolves(flowsListResponse);

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === migrateFlowsUrl) {
        return;
      }
      throw `Invalid POST request: ${opts.url}`;
    });

    await assert.doesNotReject(command.action(logger, {
      options: {
        environmentName: validEnvironment,
        solutionName: validSolutionName,
        flowName: validFlowName,
        asAdmin: true
      }
    }));

    assert.strictEqual(getDynamicsStub.calledWith(validEnvironment, true), true);
  });

  it('prompts when multiple cloud flows share the same name', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').resolves(dynamicsApiUrl);
    sinon.stub(powerPlatform, 'getSolutionByName').resolves(solutionResponse);
    sinon.stub(odata, 'getAllItems').resolves(multipleFlowsResponse);
    sinon.stub(cli, 'handleMultipleResultsFound').resolves({ name: validFlowId, properties: { displayName: validFlowName } });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === migrateFlowsUrl) {
        return;
      }
      throw `Invalid POST request: ${opts.url}`;
    });

    await assert.doesNotReject(command.action(logger, {
      options: {
        environmentName: validEnvironment,
        solutionName: validSolutionName,
        flowName: validFlowName
      }
    }));
  });

  it('throws an error when the specified cloud flow name does not exist', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').resolves(dynamicsApiUrl);
    sinon.stub(powerPlatform, 'getSolutionByName').resolves(solutionResponse);
    sinon.stub(odata, 'getAllItems').resolves([]);

    await assert.rejects(
      command.action(logger, {
        options: {
          environmentName: validEnvironment,
          solutionName: validSolutionName,
          flowName: 'NonExistentFlow'
        }
      }),
      new CommandError(`The specified flow 'NonExistentFlow' does not exist.`)
    );
  });

  it('throws an error when the specified solution name does not exist', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').resolves(dynamicsApiUrl);
    sinon.stub(powerPlatform, 'getSolutionByName').rejects(new Error(`The specified solution 'NonExistentSolution' does not exist.`));

    await assert.rejects(
      command.action(logger, {
        options: {
          environmentName: validEnvironment,
          solutionName: 'NonExistentSolution',
          flowName: validFlowName
        }
      }),
      new CommandError(`The specified solution 'NonExistentSolution' does not exist.`)
    );
  });

  it('logs verbose output during execution', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').resolves(dynamicsApiUrl);
    sinon.stub(powerPlatform, 'getSolutionByName').resolves(solutionResponse);
    sinon.stub(odata, 'getAllItems').resolves(flowsListResponse);
    sinon.stub(request, 'post').resolves();

    await command.action(logger, {
      options: {
        environmentName: validEnvironment,
        solutionName: validSolutionName,
        flowName: validFlowName,
        verbose: true
      }
    });

    assert(loggerLogToStderrSpy.called);
  });

  it('correctly handles API error', async () => {
    const errorMessage = 'Something went wrong';
    sinon.stub(request, 'post').rejects({ error: { error: { message: errorMessage } } });

    await assert.rejects(
      command.action(logger, {
        options: {
          environmentName: validEnvironment,
          solutionId: validSolutionId,
          flowId: validFlowId
        }
      }),
      new CommandError(errorMessage)
    );
  });
});
