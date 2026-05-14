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
import command from './solution-app-add.js';

describe(commands.SOLUTION_APP_ADD, () => {
  let commandInfo: CommandInfo;

  //#region Mocked Responses
  const validEnvironment = '4be50206-9576-4237-8b17-38d8aadfaa36';
  const validSolutionId = '00000001-0000-0000-0001-00000000009b';
  const validSolutionName = 'MySolution';
  const validAppId = '394378d7-c92d-4633-a7eb-0900179ee587';
  const validAppName = 'My Canvas App';
  const dynamicsApiUrl = 'https://contoso-dev.api.crm4.dynamics.com';
  const makeSolutionAwareUrl = `https://api.powerapps.com/providers/Microsoft.PowerApps/environments/${validEnvironment}/apps/${validAppId}/makeSolutionAware?api-version=2021-02-01`;

  const solutionResponse = {
    solutionid: validSolutionId,
    uniquename: validSolutionName,
    friendlyname: validSolutionName
  };

  const appsListResponse = [
    {
      name: validAppId,
      properties: { displayName: validAppName }
    }
  ];

  const multipleAppsResponse = [
    {
      name: validAppId,
      properties: { displayName: validAppName }
    },
    {
      name: 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee',
      properties: { displayName: validAppName }
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
    assert.strictEqual(command.name, commands.SOLUTION_APP_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct option sets', () => {
    const optionSets = command.optionSets;
    assert.deepStrictEqual(optionSets, [
      { options: ['solutionId', 'solutionName'] },
      { options: ['appId', 'appName'] }
    ]);
  });

  it('fails validation if solutionId is not a valid GUID', async () => {
    const actual = await command.validate({
      options: {
        environmentName: validEnvironment,
        solutionId: 'invalid-guid',
        appId: validAppId
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if appId is not a valid GUID', async () => {
    const actual = await command.validate({
      options: {
        environmentName: validEnvironment,
        solutionId: validSolutionId,
        appId: 'invalid-guid'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when solutionId and appId are valid GUIDs', async () => {
    const actual = await command.validate({
      options: {
        environmentName: validEnvironment,
        solutionId: validSolutionId,
        appId: validAppId
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when solutionName and appName are provided', async () => {
    const actual = await command.validate({
      options: {
        environmentName: validEnvironment,
        solutionName: validSolutionName,
        appName: validAppName
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('adds a canvas app to a solution using IDs and sends correct payload', async () => {
    let postBody: any;
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === makeSolutionAwareUrl) {
        postBody = opts.data;
        return;
      }
      throw `Invalid POST request: ${opts.url}`;
    });

    await assert.doesNotReject(command.action(logger, {
      options: {
        environmentName: validEnvironment,
        solutionId: validSolutionId,
        appId: validAppId
      }
    }));

    assert.deepStrictEqual(postBody, { solutionId: validSolutionId });
  });

  it('adds a canvas app to a solution by resolving names via Dataverse and PowerApps API', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').resolves(dynamicsApiUrl);
    sinon.stub(powerPlatform, 'getSolutionByName').resolves(solutionResponse);
    sinon.stub(odata, 'getAllItems').resolves(appsListResponse);

    let postBody: any;
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === makeSolutionAwareUrl) {
        postBody = opts.data;
        return;
      }
      throw `Invalid POST request: ${opts.url}`;
    });

    await assert.doesNotReject(command.action(logger, {
      options: {
        environmentName: validEnvironment,
        solutionName: validSolutionName,
        appName: validAppName
      }
    }));

    assert.deepStrictEqual(postBody, { solutionId: validSolutionId });
  });

  it('resolves canvas app name case-insensitively', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').resolves(dynamicsApiUrl);
    sinon.stub(powerPlatform, 'getSolutionByName').resolves(solutionResponse);
    sinon.stub(odata, 'getAllItems').resolves(appsListResponse);

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === makeSolutionAwareUrl) {
        return;
      }
      throw `Invalid POST request: ${opts.url}`;
    });

    await assert.doesNotReject(command.action(logger, {
      options: {
        environmentName: validEnvironment,
        solutionName: validSolutionName,
        appName: validAppName.toUpperCase()
      }
    }));
  });

  it('adds a canvas app to a solution as admin', async () => {
    const getDynamicsStub = sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').resolves(dynamicsApiUrl);
    sinon.stub(powerPlatform, 'getSolutionByName').resolves(solutionResponse);
    sinon.stub(odata, 'getAllItems').resolves(appsListResponse);

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === makeSolutionAwareUrl) {
        return;
      }
      throw `Invalid POST request: ${opts.url}`;
    });

    await assert.doesNotReject(command.action(logger, {
      options: {
        environmentName: validEnvironment,
        solutionName: validSolutionName,
        appName: validAppName,
        asAdmin: true
      }
    }));

    assert.strictEqual(getDynamicsStub.calledWith(validEnvironment, true), true);
  });

  it('prompts when multiple canvas apps share the same name', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').resolves(dynamicsApiUrl);
    sinon.stub(powerPlatform, 'getSolutionByName').resolves(solutionResponse);
    sinon.stub(odata, 'getAllItems').resolves(multipleAppsResponse);
    sinon.stub(cli, 'handleMultipleResultsFound').resolves({ name: validAppId, properties: { displayName: validAppName } });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === makeSolutionAwareUrl) {
        return;
      }
      throw `Invalid POST request: ${opts.url}`;
    });

    await assert.doesNotReject(command.action(logger, {
      options: {
        environmentName: validEnvironment,
        solutionName: validSolutionName,
        appName: validAppName
      }
    }));
  });

  it('throws an error when the specified canvas app name does not exist', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').resolves(dynamicsApiUrl);
    sinon.stub(powerPlatform, 'getSolutionByName').resolves(solutionResponse);
    sinon.stub(odata, 'getAllItems').resolves([]);

    await assert.rejects(
      command.action(logger, {
        options: {
          environmentName: validEnvironment,
          solutionName: validSolutionName,
          appName: 'NonExistentApp'
        }
      }),
      new CommandError(`The specified app 'NonExistentApp' does not exist.`)
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
          appName: validAppName
        }
      }),
      new CommandError(`The specified solution 'NonExistentSolution' does not exist.`)
    );
  });

  it('logs verbose output during execution', async () => {
    sinon.stub(powerPlatform, 'getDynamicsInstanceApiUrl').resolves(dynamicsApiUrl);
    sinon.stub(powerPlatform, 'getSolutionByName').resolves(solutionResponse);
    sinon.stub(odata, 'getAllItems').resolves(appsListResponse);
    sinon.stub(request, 'post').resolves();

    await command.action(logger, {
      options: {
        environmentName: validEnvironment,
        solutionName: validSolutionName,
        appName: validAppName,
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
          appId: validAppId
        }
      }),
      new CommandError(errorMessage)
    );
  });
});
