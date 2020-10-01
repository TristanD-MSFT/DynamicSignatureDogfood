// ================================================================================================================================================
// OUTLOOK BATCH API + BIZ LOGIC BEGINS HERE
// ================================================================================================================================================
/******/ (function(modules) { // webpackBootstrap
/******/  // The module cache
/******/  var installedModules = {};
/******/
/******/  // The require function
/******/  function __webpack_require__(moduleId) {
/******/
/******/    // Check if module is in cache
/******/    if(installedModules[moduleId]) {
/******/      return installedModules[moduleId].exports;
/******/    }
/******/    // Create a new module (and put it into the cache)
/******/    var module = installedModules[moduleId] = {
/******/      i: moduleId,
/******/      l: false,
/******/      exports: {}
/******/    };
/******/
/******/    // Execute the module function
/******/    modules[moduleId].call(module.exports, module, module.exports, __webpack_require__);
/******/
/******/    // Flag the module as loaded
/******/    module.l = true;
/******/
/******/    // Return the exports of the module
/******/    return module.exports;
/******/  }
/******/
/******/
/******/  // expose the modules object (__webpack_modules__)
/******/  __webpack_require__.m = modules;
/******/
/******/  // expose the module cache
/******/  __webpack_require__.c = installedModules;
/******/
/******/  // define getter function for harmony exports
/******/  __webpack_require__.d = function(exports, name, getter) {
/******/    if(!__webpack_require__.o(exports, name)) {
/******/      Object.defineProperty(exports, name, { enumerable: true, get: getter });
/******/    }
/******/  };
/******/
/******/  // define __esModule on exports
/******/  __webpack_require__.r = function(exports) {
/******/    if(typeof Symbol !== 'undefined' && Symbol.toStringTag) {
/******/      Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/    }
/******/    Object.defineProperty(exports, '__esModule', { value: true });
/******/  };
/******/
/******/  // create a fake namespace object
/******/  // mode & 1: value is a module id, require it
/******/  // mode & 2: merge all properties of value into the ns
/******/  // mode & 4: return value when already ns object
/******/  // mode & 8|1: behave like require
/******/  __webpack_require__.t = function(value, mode) {
/******/    if(mode & 1) value = __webpack_require__(value);
/******/    if(mode & 8) return value;
/******/    if((mode & 4) && typeof value === 'object' && value && value.__esModule) return value;
/******/    var ns = Object.create(null);
/******/    __webpack_require__.r(ns);
/******/    Object.defineProperty(ns, 'default', { enumerable: true, value: value });
/******/    if(mode & 2 && typeof value != 'string') for(var key in value) __webpack_require__.d(ns, key, function(key) { return value[key]; }.bind(null, key));
/******/    return ns;
/******/  };
/******/
/******/  // getDefaultExport function for compatibility with non-harmony modules
/******/  __webpack_require__.n = function(module) {
/******/    var getter = module && module.__esModule ?
/******/      function getDefault() { return module['default']; } :
/******/      function getModuleExports() { return module; };
/******/    __webpack_require__.d(getter, 'a', getter);
/******/    return getter;
/******/  };
/******/
/******/  // Object.prototype.hasOwnProperty.call
/******/  __webpack_require__.o = function(object, property) { return Object.prototype.hasOwnProperty.call(object, property); };
/******/
/******/  // __webpack_public_path__
/******/  __webpack_require__.p = "";
/******/
/******/
/******/  // Load entry module and return exports
/******/  return __webpack_require__(__webpack_require__.s = 0);
/******/ })
/************************************************************************/
/******/ ([
/* 0 */
/***/ (function(module, exports, __webpack_require__) {

  "use strict";

  var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
      return new (P || (P = Promise))(function (resolve, reject) {
          function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
          function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
          function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
          step((generator = generator.apply(thisArg, _arguments || [])).next());
      });
  };
  var __generator = (this && this.__generator) || function (thisArg, body) {
      var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
      return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
      function verb(n) { return function (v) { return step([n, v]); }; }
      function step(op) {
          if (f) throw new TypeError("Generator is already executing.");
          while (_) try {
              if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
              if (y = 0, t) op = [op[0] & 2, t.value];
              switch (op[0]) {
                  case 0: case 1: t = op; break;
                  case 4: _.label++; return { value: op[1], done: false };
                  case 5: _.label++; y = op[1]; op = [0]; continue;
                  case 7: op = _.ops.pop(); _.trys.pop(); continue;
                  default:
                      if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                      if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                      if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                      if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                      if (t[2]) _.ops.pop();
                      _.trys.pop(); continue;
              }
              op = body.call(thisArg, _);
          } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
          if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
      }
  };
  Object.defineProperty(exports, "__esModule", { value: true });
  var Outlook = __webpack_require__(1);
  var GlobalInvocationId = 0;
  ;
  var GlobalMailboxCallbackTable;
  var GlobalMailboxApiAsyncContextTable;
  function invokeMailboxCreateOM(initialData) {
      console.log("invokeMailboxCreateOM was invoked!" + initialData);
      MailboxCreateOM(initialData);
  }
  function InvokeMailboxCreateOM(initialData) {
      console.log("InvokeMailboxCreateOM was invoked!" + initialData);
      MailboxCreateOM(initialData);
  }
  function SendOmReadyMessageToHost() {
      return __awaiter(this, void 0, void 0, function () {
          var context;
          return __generator(this, function (_a) {
              switch (_a.label) {
                  case 0:
                      context = new Outlook.RequestContext();
                      context.mailbox.onSdxControlReadyResponse();
                      return [4 /*yield*/, context.sync()];
                  case 1:
                      _a.sent();
                      console.log("SDX Control is ready!");
                      return [2 /*return*/];
              }
          });
      });
  }
  exports.SendOmReadyMessageToHost = SendOmReadyMessageToHost;
  function MailboxHostExecuteApi(dispid, asyncContext, callback, parameters) {
      return __awaiter(this, void 0, void 0, function () {
          var invocationId, context;
          return __generator(this, function (_a) {
              if (!GlobalMailboxCallbackTable) {
                  GlobalMailboxCallbackTable = {};
              }
              if (!GlobalMailboxApiAsyncContextTable) {
                  GlobalMailboxApiAsyncContextTable = {};
              }
              invocationId = ++GlobalInvocationId;
              if (callback) {
                  GlobalMailboxCallbackTable[invocationId] = callback;
              }
              if (asyncContext) {
                  GlobalMailboxApiAsyncContextTable[invocationId] = asyncContext;
              }
              context = new Outlook.RequestContext();
              context.mailbox.executeApi(dispid, invocationId, parameters);
              context.sync();
              console.log("MailboxHostExecuteApi invoked!");
              return [2 /*return*/];
          });
      });
  }
  function onNewComposeHandler(eventObj)
  {

    //get current time
	  let today = new Date();
	  let time = today.getHours();
	  let day = today.getDay();
  
  	if (day == 0 || day == 6 || time < 8 || time > 16) {
      Office.context.mailbox.item.body.setSignatureAsync(
        "<p style='margin-bottom:0in;line-height:normal'><span style='font-size:9.0pt'>------------</span></p>" + 
        "<p style='margin-bottom:0in;line-height:normal'><span style='font-size:9.0pt'>Your family and personal time is important to me; after-hours responses not required or expected!</span></p >",
        { coercionType: Office.CoercionType.Html }
      );
    }
    else{
      Office.context.mailbox.item.body.setSignatureAsync(
        "<p style='margin-bottom:0in;line-height:normal'><span style='font-size:9.0pt'>------------</span></p>" + 
        "<p style='margin-bottom:0in;line-height:normal'><span style='font-size:9.0pt'>ANSWER ASAP PLEASE!</span></p >",
        { coercionType: Office.CoercionType.Html }
      );
    }
    eventObj.completed();
  }
  exports.onNewComposeHandler = onNewComposeHandler;
  var LaunchEventObj = /** @class */ (function () {
      function LaunchEventObj(eventType) {
          this.launchEventType = eventType;
      }
      LaunchEventObj.prototype.getEventType = function () {
          return this.launchEventType;
      };
      LaunchEventObj.prototype.completed = function (data) {
          CustomFunctionMappings['SendEventCompletedMessageToHost'](JSON.stringify(data), this.launchEventType);
      };
      return LaunchEventObj;
  }());
  var GlobalLaunchEventObj;
  function createEventObj(launchEventType) {
      if (!GlobalLaunchEventObj) {
          GlobalLaunchEventObj = new LaunchEventObj(launchEventType);
      }
      return GlobalLaunchEventObj;
  }
  function RunLaunchEventHandlerFunction(funcName, launchEventType) {
      return __awaiter(this, void 0, void 0, function () {
          var eventType;
          return __generator(this, function (_a) {
              eventType = +launchEventType;
              CustomFunctionMappings[funcName](createEventObj(eventType));
              console.log("RunLaunchEventHandlerFunction was invoked with funcName = " + funcName);
              return [2 /*return*/];
          });
      });
  }
  exports.RunLaunchEventHandlerFunction = RunLaunchEventHandlerFunction;
  function SendEventCompletedMessageToHost(data, launchEventType) {
      return __awaiter(this, void 0, void 0, function () {
          var context;
          return __generator(this, function (_a) {
              switch (_a.label) {
                  case 0:
                      context = new Outlook.RequestContext();
                      context.mailbox.onLaunchEventCompleted(data, launchEventType);
                      return [4 /*yield*/, context.sync()];
                  case 1:
                      _a.sent();
                      console.log("Event completed message sent to host!");
                      return [2 /*return*/];
              }
          });
      });
  }
  exports.SendEventCompletedMessageToHost = SendEventCompletedMessageToHost;
  function MailboxApiCallComplete(data, invocationId) {
      return __awaiter(this, void 0, void 0, function () {
          var mailboxApiCallback, mailboxApiAsyncContext, responseData;
          return __generator(this, function (_a) {
              mailboxApiCallback = GlobalMailboxCallbackTable[invocationId];
              mailboxApiAsyncContext = GlobalMailboxApiAsyncContextTable[invocationId];
              if (!mailboxApiCallback) {
                  console.log("no api callback found for invocationId: " + invocationId);
              }
              else {
                  // responseData = JSON.parse(data);
                  // mailboxApiCallback({ "asyncContext": mailboxApiAsyncContext, "status": "succeeded" });

                  responseData = JSON.parse(data);
                  let results;
                  if(responseData){
                    results = JSON.parse(responseData[0]);
                  }
                  if(results && results.data != null && results.wasSuccessful != null){
                    if(results.wasSuccessful){
                      mailboxApiCallback({ "asyncContext": mailboxApiAsyncContext, "status": "succeeded", "value": results.data});
                    }
                    else{
                      mailboxApiCallback({ "asyncContext": mailboxApiAsyncContext, "status": "failed" });
                    }
                   
                  }
                  else{
                    mailboxApiCallback({ "asyncContext": mailboxApiAsyncContext, "status": "succeeded"});
                  }

                  // let x = 6;
                  // ++x;
                  // let asyncResult = CustomFunctionMappings['launchEventCreateAsyncResult'](data);
                  // mailboxApiCallback(asyncResult);
                  // delete mailboxApiCallback[invocationId];
              }
              return [2 /*return*/];
          });
      });
  }
  exports.MailboxApiCallComplete = MailboxApiCallComplete;
  
  // Global functions that can be invoked by the host or the add-in.
  CustomFunctionMappings['INVOKE_MAILBOX_CREATE_OM'] = InvokeMailboxCreateOM;
  CustomFunctionMappings['RUN_LAUNCH_EVENT_HANDLER_FUNCTION'] = RunLaunchEventHandlerFunction;
  CustomFunctionMappings['MAILBOX_API_CALL_COMPLETE'] = MailboxApiCallComplete;
  CustomFunctionMappings['SendOmReadyMessageToHost'] = SendOmReadyMessageToHost;
  CustomFunctionMappings['MailboxHostExecuteApi'] = MailboxHostExecuteApi;
  CustomFunctionMappings['SendEventCompletedMessageToHost'] = SendEventCompletedMessageToHost;

  CustomFunctionMappings['onNewComposeHandler'] = onNewComposeHandler;
  
  /***/ }),
  /* 1 */
  /***/ (function(module, exports, __webpack_require__) {
  
  "use strict";
  
  var __extends = (this && this.__extends) || (function () {
      var extendStatics = function (d, b) {
          extendStatics = Object.setPrototypeOf ||
              ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
              function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
          return extendStatics(d, b);
      }
      return function (d, b) {
          extendStatics(d, b);
          function __() { this.constructor = d; }
          d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
      };
  })();
  Object.defineProperty(exports, "__esModule", { value: true });
  /*
   * This is a generated file.  Generated by outlook\outllib\webext\api\codegen\OutlookApiGen.bat.
   * If there are content placeholders, only edit content inside content placeholders.
   * If there are no content placeholders, do not edit this file directly.
   */
  /* Begin_PlaceHolder_GlobalHeader */
  //
  // NOTE : =====================================================================================================================================================================================================================================
  //  Make the following changes to this file after moving it to fabric-internal before invoking npm run build:
  //  1 - Uncomment the imports below.
  //  2 - Remove the Outlook namespace.
  //  3 - Search for "128e855c-bdb5-4fbc-b2ac-b37b29b0b1a4" and then un-comment the block of code right after it.
  // NOTE : =====================================================================================================================================================================================================================================
  //
  var OfficeExtension = __webpack_require__(2);
  // import * as OfficeCore from '@microsoft/office-core-batch';
  /* End_PlaceHolder_GlobalHeader */
  /* Begin_PlaceHolder_ModuleHeader */
  /* End_PlaceHolder_ModuleHeader */
  var _hostName = "Outlook";
  var _defaultApiSetName = "OutlookApi";
  var _createPropertyObject = OfficeExtension.BatchApiHelper.createPropertyObject;
  var _createMethodObject = OfficeExtension.BatchApiHelper.createMethodObject;
  var _createIndexerObject = OfficeExtension.BatchApiHelper.createIndexerObject;
  var _createRootServiceObject = OfficeExtension.BatchApiHelper.createRootServiceObject;
  var _createTopLevelServiceObject = OfficeExtension.BatchApiHelper.createTopLevelServiceObject;
  var _createChildItemObject = OfficeExtension.BatchApiHelper.createChildItemObject;
  var _invokeMethod = OfficeExtension.BatchApiHelper.invokeMethod;
  var _invokeEnsureUnchanged = OfficeExtension.BatchApiHelper.invokeEnsureUnchanged;
  var _invokeSetProperty = OfficeExtension.BatchApiHelper.invokeSetProperty;
  var _isNullOrUndefined = OfficeExtension.Utility.isNullOrUndefined;
  var _isUndefined = OfficeExtension.Utility.isUndefined;
  var _throwIfNotLoaded = OfficeExtension.Utility.throwIfNotLoaded;
  var _throwIfApiNotSupported = OfficeExtension.Utility.throwIfApiNotSupported;
  var _load = OfficeExtension.Utility.load;
  var _retrieve = OfficeExtension.Utility.retrieve;
  var _toJson = OfficeExtension.Utility.toJson;
  var _fixObjectPathIfNecessary = OfficeExtension.Utility.fixObjectPathIfNecessary;
  var _handleNavigationPropertyResults = OfficeExtension.Utility._handleNavigationPropertyResults;
  var _adjustToDateTime = OfficeExtension.Utility.adjustToDateTime;
  var _processRetrieveResult = OfficeExtension.Utility.processRetrieveResult;
  var _setMockData = OfficeExtension.Utility.setMockData;
  var _typeMailBox = "MailBox";
  /* Begin_PlaceHolder_MailBox_BeforeDeclaration */
  /* End_PlaceHolder_MailBox_BeforeDeclaration */
  /**
   * [Api set: OutlookApi 1.1]
   */
  var MailBox = /** @class */ (function (_super) {
      __extends(MailBox, _super);
      function MailBox() {
          return _super !== null && _super.apply(this, arguments) || this;
      }
      Object.defineProperty(MailBox.prototype, "_className", {
          get: function () {
              return "MailBox";
          },
          enumerable: true,
          configurable: true
      });
      /**
       * The shape is used so that it should be a very close replacement to the shape needed by OutlookAppOm CallOutlookDispatcher
       *
       * [Api set: OutlookApi 1.1]
       *
       * @param dispid API dispid
       * @param invocationId Unique ID that is passed down from Outlook JS for each API call.
       * @param parameters array of parameters
       * @returns void
       */
      MailBox.prototype.executeApi = function (dispid, invocationId, parameters) {
          /* Begin_PlaceHolder_MailBox_ExecuteApi */
          /* End_PlaceHolder_MailBox_ExecuteApi */
          _invokeMethod(this, "ExecuteApi", 0 /* Default */, [dispid, invocationId, parameters], 0 /* none */, 0 /* none */);
      };
      /**
       * Notifies the host that the LaunchEvent add-in is done executing and can be closed.
       *
       * [Api set: OutlookApi 1.1]
       * @returns void
       */
      MailBox.prototype.onLaunchEventCompleted = function (jsonData, launchEventType) {
          /* Begin_PlaceHolder_MailBox_OnLaunchEventCompleted */
          /* End_PlaceHolder_MailBox_OnLaunchEventCompleted */
          _invokeMethod(this, "OnLaunchEventCompleted", 0 /* Default */, [jsonData, launchEventType], 0 /* none */, 0 /* none */);
      };
      /**
       * Notifies the host after the Mailbox OM has been created/initialized in the react native sandbox.
                  This means that the add-in instance is running and ready for back and forth communication with the host.
       *
       * [Api set: OutlookApi 1.1]
       * @returns void
       */
      MailBox.prototype.onSdxControlReadyResponse = function () {
          /* Begin_PlaceHolder_MailBox_OnSdxControlReadyResponse */
          /* End_PlaceHolder_MailBox_OnSdxControlReadyResponse */
          _invokeMethod(this, "OnSdxControlReadyResponse", 0 /* Default */, [], 0 /* none */, 0 /* none */);
      };
      /** Handle results returned from the document
       * @private
       */
      MailBox.prototype._handleResult = function (value) {
          _super.prototype._handleResult.call(this, value);
          if (_isNullOrUndefined(value))
              return;
          var obj = value;
          _fixObjectPathIfNecessary(this, obj);
          /* Begin_PlaceHolder_MailBox_HandleResult */
          /* End_PlaceHolder_MailBox_HandleResult */
      };
      /** Handle retrieve results
       * @private
       */
      MailBox.prototype._handleRetrieveResult = function (value, result) {
          _super.prototype._handleRetrieveResult.call(this, value, result);
          /* Begin_PlaceHolder_MailBox_HandleRetrieveResult */
          /* End_PlaceHolder_MailBox_HandleRetrieveResult */
          _processRetrieveResult(this, value, result);
      };
      /**
      * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
      * Whereas the original Outlook.MailBox object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Outlook.Interfaces.MailBoxData`) that contains shallow copies of any loaded child properties from the original object.
      */
      MailBox.prototype.toJSON = function () {
          return _toJson(this, 
          /* scalarProperties: */ {}, 
          /* navigationProperties: */ {});
      };
      return MailBox;
  }(OfficeExtension.ClientObject));
  exports.MailBox = MailBox;
  /* Begin_PlaceHolder_MailBox_AfterDeclaration */
  // Block 128e855c-bdb5-4fbc-b2ac-b37b29b0b1a4
  var RequestContext = /** @class */ (function (_super) {
      __extends(RequestContext, _super);
      function RequestContext() {
          var _this = _super.call(this) || this;
          _this.m_mailbox = _createRootServiceObject(MailBox, _this);
          return _this;
      }
      Object.defineProperty(RequestContext.prototype, "mailbox", {
          get: function () {
              return this.m_mailbox;
          },
          enumerable: true,
          configurable: true
      });
      return RequestContext;
  }(OfficeExtension.ClientRequestContext));
  exports.RequestContext = RequestContext;
  /* Begin_PlaceHolder_ErrorCodesTypeName */
  var ErrorCodes;
  (function (ErrorCodes) {
      ErrorCodes["generalException"] = "GeneralException";
      /* Begin_PlaceHolder_ErrorCodesAdditional */
      /* End_PlaceHolder_ErrorCodesAdditional */
  })(ErrorCodes = exports.ErrorCodes || (exports.ErrorCodes = {}));
  /* Begin_PlaceHolder_GlobalFooter */
  /* End_PlaceHolder_GlobalFooter */
  
  
  /***/ }),
  /* 2 */
  /***/ (function(module, exports) {
  
  module.exports = OfficeExtensionBatch;
  
  /***/ })
  /******/ ]);
  //# sourceMappingURL=bundle.js.map
// ================================================================================================================================================
// OUTLOOK BATCH API + BIZ LOGIC ENDS HERE
// ================================================================================================================================================



// ================================================================================================================================================
// OUTLOOKJS CODE STARTS HERE
// ================================================================================================================================================
/* Outlook LaunchEvent API library */
/* osf version: 16.0.12409.10000 */
/* office-js-api version: 0.0.0.0 */
/*
	Copyright (c) Microsoft Corporation.  All rights reserved.
*/
/*
    Your use of this file is governed by the Microsoft Services Agreement http://go.microsoft.com/fwlink/?LinkId=266419.

    This file also contains the following Promise implementation (with a few small modifications):
        * @overview es6-promise - a tiny implementation of Promises/A+.
        * @copyright Copyright (c) 2014 Yehuda Katz, Tom Dale, Stefan Penner and contributors (Conversion to ES6 API by Jake Archibald)
        * @license   Licensed under MIT license
        *            See https://raw.githubusercontent.com/jakearchibald/es6-promise/master/LICENSE
        * @version   2.3.0
*/
Strings = typeof Strings === "object" ? Strings : {};
Strings.OfficeOM = Strings.OfficeOM || {};

Strings.OfficeOM.L_APICallFailed='API Call Failed';
Strings.OfficeOM.L_APINotSupported='API Not Supported';
Strings.OfficeOM.L_ActivityLimitReached='Activity limit has been reached.';
Strings.OfficeOM.L_AddBindingFromPromptDefaultText='Please make a selection.';
Strings.OfficeOM.L_AddinIsAlreadyRequestingToken='Add-in is already requesting an access token.';
Strings.OfficeOM.L_AddinIsAlreadyRequestingTokenMessage='The operation failed because this add-in is already requesting an access token.';
Strings.OfficeOM.L_ApiNotFoundDetails='The method or property {0} is part of the {1} requirement set, which is not available in your version of {2}.';
Strings.OfficeOM.L_AppNameNotExist='Add-in Name for {0} doesn\'t exist.';
Strings.OfficeOM.L_AppNotExistInitializeNotCalled='Application {0} does not exist. Microsoft.Office.WebExtension.initialize(reason) is not called.';
Strings.OfficeOM.L_AttemptingToSetReadOnlyProperty='Attempting to set read-only property \'{0}\'.';
Strings.OfficeOM.L_BadSelectorString='The string passed into the selector is improperly formatted or unsupported.';
Strings.OfficeOM.L_BindingCreationError='Binding Creation Error';
Strings.OfficeOM.L_BindingNotExist='The specified binding does not exist.';
Strings.OfficeOM.L_BindingToMultipleSelection='Noncontiguous selections are not supported.';
Strings.OfficeOM.L_BrowserAPINotSupported='This browser does not support the requested API.';
Strings.OfficeOM.L_CallbackNotAFunction='Callback must be of type function, was of type {0}.';
Strings.OfficeOM.L_CannotApplyPropertyThroughSetMethod='Changes to property \'{0}\' cannot be applied through an \"object.set\" method.';
Strings.OfficeOM.L_CannotNavigateTo='The object is located in a place where navigation is not supported.';
Strings.OfficeOM.L_CannotRegisterEvent='The event handler cannot be registered.';
Strings.OfficeOM.L_CannotWriteToSelection='Cannot write to the current selection.';
Strings.OfficeOM.L_CellDataAmountBeyondLimits='Note: The number of cells in a table is suggested to be below 20,000 cells.';
Strings.OfficeOM.L_CellFormatAmountBeyondLimits='Note: The formatting sets set by a Formatting API call is suggested to be below 100.';
Strings.OfficeOM.L_CloseFileBeforeRetrieve='Call closeAsync on the current File before retrieving another.';
Strings.OfficeOM.L_CoercionTypeNotMatchBinding='The specified coercion type is not compatible with this binding type.';
Strings.OfficeOM.L_CoercionTypeNotSupported='The specified coercion type is not supported.';
Strings.OfficeOM.L_ColIndexOutOfRange='The column index value is out of the allowed range. Use a value (0 or higher) that\'s less than the number of columns.';
Strings.OfficeOM.L_ConnectionFailureWithDetails='The request failed with status code of {0}, error code {1} and the following error message: {2}';
Strings.OfficeOM.L_ConnectionFailureWithStatus='The request failed with status code of {0}.';
Strings.OfficeOM.L_CustomFunctionDefinitionMissing='A property with this name that represents the function\'s definition must exist on Excel.Script.CustomFunctions.';
Strings.OfficeOM.L_CustomFunctionImplementationMissing='The property with this name on Excel.Script.CustomFunctions that represents the function\'s definition must contain a \'call\' property that implements the function.';
Strings.OfficeOM.L_CustomFunctionNameCannotSplit='The function name must contain a non-empty namespace and a non-empty short name.';
Strings.OfficeOM.L_CustomFunctionNameContainsBadChars='The function name may only contain letters, digits, underscores, and periods.';
Strings.OfficeOM.L_CustomXmlError='Custom XML Error.';
Strings.OfficeOM.L_CustomXmlExceedQuotaMessage='XPath limits selection to 1024 items.';
Strings.OfficeOM.L_CustomXmlExceedQuotaName='Selection Limit Reached';
Strings.OfficeOM.L_CustomXmlNodeNotFound='The specified node was not found.';
Strings.OfficeOM.L_CustomXmlOutOfDateMessage='The data is out of date. Retrieve the object again.';
Strings.OfficeOM.L_CustomXmlOutOfDateName='Data Not Current';
Strings.OfficeOM.L_DataNotMatchBindingSize='The supplied data object does not match the size of the current selection.';
Strings.OfficeOM.L_DataNotMatchBindingType='The specified data object is not compatible with the binding type.';
Strings.OfficeOM.L_DataNotMatchCoercionType='The type of the specified data object is not compatible with the current selection.';
Strings.OfficeOM.L_DataNotMatchSelection='The supplied data object is not compatible with the shape or dimensions of the current selection.';
Strings.OfficeOM.L_DataReadError='Data Read Error';
Strings.OfficeOM.L_DataStale='Data Not Current';
Strings.OfficeOM.L_DataWriteError='Data Write Error';
Strings.OfficeOM.L_DataWriteReminder='Data Write Reminder';
Strings.OfficeOM.L_DialogAddressNotTrusted='The domain of the URL is not included in the AppDomains element in the manifest, and is not subdomain of source location.';
Strings.OfficeOM.L_DialogAlreadyOpened='The operation failed because this add-in already has an active dialog.';
Strings.OfficeOM.L_DialogInvalidScheme='The URL scheme is not supported. Use HTTPS instead.';
Strings.OfficeOM.L_DialogNavigateError='Dialog Navigation Error';
Strings.OfficeOM.L_DialogOK='OK';
Strings.OfficeOM.L_DialogRequireHTTPS='The HTTP protocol is not supported. Use HTTPS instead';
Strings.OfficeOM.L_DisplayDialogError='Display Dialog Error';
Strings.OfficeOM.L_DocumentReadOnly='The requested operation is not allowed on the current document mode.';
Strings.OfficeOM.L_ElementMissing='We couldn\'t format the table cell because some parameter values are missing. Double-check the parameters and try again.';
Strings.OfficeOM.L_EventHandlerAdditionFailed='Failed to add the event handler.';
Strings.OfficeOM.L_EventHandlerNotExist='The specified event handler was not found for this binding.';
Strings.OfficeOM.L_EventHandlerRemovalFailed='Failed to remove the event handler.';
Strings.OfficeOM.L_EventRegistrationError='Event Registration Error';
Strings.OfficeOM.L_FileTypeNotSupported='The specified file type is not supported.';
Strings.OfficeOM.L_FormatValueOutOfRange='The value is out of the allowed range.';
Strings.OfficeOM.L_FormattingReminder='Formatting Reminder';
Strings.OfficeOM.L_FunctionCallFailed='Function {0} call failed, error code: {1}.';
Strings.OfficeOM.L_GetDataIsTooLarge='The requested data set is too large.';
Strings.OfficeOM.L_GetDataParametersConflict='The specified parameters conflict.';
Strings.OfficeOM.L_GetSelectionNotSupported='The current selection is not supported.';
Strings.OfficeOM.L_HostError='Host Error';
Strings.OfficeOM.L_ImplicitGetAuthContextMissing='The function to get authentication context is missing';
Strings.OfficeOM.L_ImplicitNotLoaded='The module is not loaded before acquire a token';
Strings.OfficeOM.L_InValidOptionalArgument='invalid optional argument';
Strings.OfficeOM.L_IndexOutOfRange='Index out of range.';
Strings.OfficeOM.L_InitializeNotReady='Office.js has not been fully loaded yet. Please try again later or make sure to add your initialization code on the Office.initialize function.';
Strings.OfficeOM.L_InternalError='Internal Error';
Strings.OfficeOM.L_InternalErrorDescription='An internal error has occurred.';
Strings.OfficeOM.L_InvalidAPICall='Invalid API Call';
Strings.OfficeOM.L_InvalidApiArgumentsMessage='Invalid input arguments.';
Strings.OfficeOM.L_InvalidApiCallInContext='Invalid API call in the current context.';
Strings.OfficeOM.L_InvalidArgument='The argument \'{0}\' doesn\'t work for this situation, is missing, or isn\'t in the right format.';
Strings.OfficeOM.L_InvalidArgumentGeneric='The argument(s) passed into the function don\'t work for this situation, are missing, or aren\'t in the right format.';
Strings.OfficeOM.L_InvalidBinding='Invalid Binding';
Strings.OfficeOM.L_InvalidBindingError='Invalid Binding Error';
Strings.OfficeOM.L_InvalidBindingOperation='Invalid Binding Operation';
Strings.OfficeOM.L_InvalidCellsValue='One or more of the cells parameters have values that aren\'t allowed. Double-check the values and try again.';
Strings.OfficeOM.L_InvalidCoercion='Invalid Coercion Type';
Strings.OfficeOM.L_InvalidColumnsForBinding='The specified columns are invalid.';
Strings.OfficeOM.L_InvalidDataFormat='The format of the specified data object is invalid.';
Strings.OfficeOM.L_InvalidDataObject='Invalid Data Object';
Strings.OfficeOM.L_InvalidFormat='Invalid Format Error';
Strings.OfficeOM.L_InvalidFormatValue='One or more of the format parameters have values that aren\'t allowed. Double-check the values and try again.';
Strings.OfficeOM.L_InvalidGetColumns='The specified columns are invalid.';
Strings.OfficeOM.L_InvalidGetRowColumnCounts='The specified rowCount or columnCount values are invalid.';
Strings.OfficeOM.L_InvalidGetRows='The specified rows are invalid.';
Strings.OfficeOM.L_InvalidGetStartRowColumn='The specified startRow or startColumn values are invalid.';
Strings.OfficeOM.L_InvalidGrant='Preauthorization missing.';
Strings.OfficeOM.L_InvalidGrantMessage='Missing grant for this add-in.';
Strings.OfficeOM.L_InvalidNamedItemForBindingType='The specified binding type is not compatible with the supplied named item.';
Strings.OfficeOM.L_InvalidNode='Invalid Node';
Strings.OfficeOM.L_InvalidObjectPath='The object path \'{0}\' isn\'t working for what you\'re trying to do. If you\'re using the object across multiple \"context.sync\" calls and outside the sequential execution of a \".run\" batch, please use the \"context.trackedObjects.add()\" and \"context.trackedObjects.remove()\" methods to manage the object\'s lifetime.';
Strings.OfficeOM.L_InvalidOperationInCellEditMode='Excel is in cell-editing mode. Please exit the edit mode by pressing ENTER or TAB or selecting another cell, and then try again.';
Strings.OfficeOM.L_InvalidOrTimedOutSession='Invalid Or Timed Out Session';
Strings.OfficeOM.L_InvalidOrTimedOutSessionMessage='Your Office session has expired or is invalid. To continue, refresh the page.';
Strings.OfficeOM.L_InvalidParameters='Function {0} has invalid parameters.';
Strings.OfficeOM.L_InvalidReadForBlankRow='The specified row is blank.';
Strings.OfficeOM.L_InvalidRequestContext='Cannot use the object across different request contexts.';
Strings.OfficeOM.L_InvalidResourceUrl='Invalid application resource Url provided.';
Strings.OfficeOM.L_InvalidResourceUrlMessage='Invalid resource Url specified in the manifest.';
Strings.OfficeOM.L_InvalidSSOAddinMessage='The identity API is not supported for this add-in.';
Strings.OfficeOM.L_InvalidSelectionForBindingType='A binding cannot be created with the current selection and the specified binding type.';
Strings.OfficeOM.L_InvalidSetColumns='The specified columns are invalid.';
Strings.OfficeOM.L_InvalidSetRows='The specified rows are invalid.';
Strings.OfficeOM.L_InvalidSetStartRowColumn='The specified startRow or startColumn values are invalid.';
Strings.OfficeOM.L_InvalidTableOptionValue='One or more of the tableOptions parameters have values that aren\'t allowed. Double-check the values and try again.';
Strings.OfficeOM.L_InvalidValue='Invalid Value';
Strings.OfficeOM.L_MemoryLimit='Memory Limit Exceeded';
Strings.OfficeOM.L_MissingParameter='Missing Parameter';
Strings.OfficeOM.L_MissingRequiredArguments='missing some required arguments';
Strings.OfficeOM.L_MultipleNamedItemFound='Multiple objects with the same name were found.';
Strings.OfficeOM.L_NamedItemNotFound='The named item does not exist.';
Strings.OfficeOM.L_NavOutOfBound='The operation failed because the index is out of range.';
Strings.OfficeOM.L_NetworkProblem='Network Problem';
Strings.OfficeOM.L_NetworkProblemRetrieveFile='A network problem has prevented retrieval of the file.';
Strings.OfficeOM.L_NewWindowCrossZone='The security settings in your browser prevent us from creating a dialog box. Try a different browser, or {0} so that \'{1}\' and the domain shown in your address bar are in the same security zone.';
Strings.OfficeOM.L_NewWindowCrossZoneConfigureBrowserLink='configure your browser';
Strings.OfficeOM.L_NewWindowCrossZoneErrorString='Browser restrictions prevented us from creating the dialog box. The domain of dialog box and the domain of the add-in host are not in the same security zone.';
Strings.OfficeOM.L_NoCapability='You don\'t have sufficient permissions for this action.';
Strings.OfficeOM.L_NonUniformPartialGetNotSupported='Coordinate parameters cannot be used with coercion type Table when the table contains merged cells.';
Strings.OfficeOM.L_NonUniformPartialSetNotSupported='Coordinate parameters cannot be used with coercion type Table when the table contains merged cells.';
Strings.OfficeOM.L_NotImplemented='Function {0} is not implemented.';
Strings.OfficeOM.L_NotSupported='Function {0} is not supported.';
Strings.OfficeOM.L_NotSupportedBindingType='The specified binding type {0} is not supported.';
Strings.OfficeOM.L_NotSupportedEventType='The specified event type {0} is not supported.';
Strings.OfficeOM.L_OperationCancelledError='Operation Cancelled';
Strings.OfficeOM.L_OperationCancelledErrorMessage='The operation was cancelled by the user.';
Strings.OfficeOM.L_OperationNotSupported='The operation is not supported.';
Strings.OfficeOM.L_OperationNotSupportedOnMatrixData='Selected content needs to be in table format. Format the data as a table and try again.';
Strings.OfficeOM.L_OperationNotSupportedOnThisBindingType='Operation is not supported on this binding type.';
Strings.OfficeOM.L_OsfControlTypeNotSupported='OsfControl type not supported.';
Strings.OfficeOM.L_OutOfRange='Out of Range';
Strings.OfficeOM.L_OverwriteWorksheetData='The set operation failed because the supplied data object will overwrite or shift data.';
Strings.OfficeOM.L_PermissionDenied='Permission Denied';
Strings.OfficeOM.L_PropertyDoesNotExist='Property \'{0}\' does not exist on the object.';
Strings.OfficeOM.L_PropertyNotLoaded='The property \'{0}\' is not available. Before reading the property\'s value, call the load method on the containing object and call \"context.sync()\" on the associated request context.';
Strings.OfficeOM.L_ReadSettingsError='Read Settings Error';
Strings.OfficeOM.L_RedundantCallbackSpecification='Callback cannot be specified both in argument list and in optional object.';
Strings.OfficeOM.L_RequestTimeout='The call took too long to execute.';
Strings.OfficeOM.L_RequestTokenUnavailable='This API has been throttled to slow the call frequency.';
Strings.OfficeOM.L_RowIndexOutOfRange='The row index value is out of the allowed range. Use a value (0 or higher) that\'s less than the number of rows.';
Strings.OfficeOM.L_RunMustReturnPromise='The batch function passed to the \".run\" method didn\'t return a promise. The function must return a promise, so that any automatically-tracked objects can be released at the completion of the batch operation. Typically, you return a promise by returning the response from \"context.sync()\".';
Strings.OfficeOM.L_SSOClientError='Error occurred in the authentication request from Office.';
Strings.OfficeOM.L_SSOClientErrorMessage='An unexpected error occurred in the client.';
Strings.OfficeOM.L_SSOConnectionLostError='A connection was lost during the sign in process.';
Strings.OfficeOM.L_SSOConnectionLostErrorMessage='A connection was lost during the sign in process, and the user may not be signed in. This was likely due to the user\'s browser configuration settings, such as security zones.';
Strings.OfficeOM.L_SSOServerError='Error occurred in the authentication provider.';
Strings.OfficeOM.L_SSOServerErrorMessage='An unexpected error occurred on the server.';
Strings.OfficeOM.L_SSOUnsupportedPlatform='API is not supported in this platform.';
Strings.OfficeOM.L_SSOUserConsentNotSupportedByCurrentAddinCategory='This add-in does not support user consent.';
Strings.OfficeOM.L_SSOUserConsentNotSupportedByCurrentAddinCategoryMessage='The operation failed because this add-in does not support user consent in this category';
Strings.OfficeOM.L_SaveSettingsError='Save Settings Error';
Strings.OfficeOM.L_SelectionCannotBound='Cannot bind to the current selection.';
Strings.OfficeOM.L_SelectionNotSupportCoercionType='The current selection is not compatible with the specified coercion type.';
Strings.OfficeOM.L_SetDataIsTooLarge='The specified data object is too large.';
Strings.OfficeOM.L_SetDataParametersConflict='The specified parameters conflict.';
Strings.OfficeOM.L_SettingNameNotExist='The specified setting name does not exist.';
Strings.OfficeOM.L_SettingsAreStale='Settings could not be saved because they are not current.';
Strings.OfficeOM.L_SettingsCannotSave='The settings could not be saved.';
Strings.OfficeOM.L_SettingsStaleError='Settings Stale Error';
Strings.OfficeOM.L_ShowWindowDialogNotification='{0} wants to display a new window.';
Strings.OfficeOM.L_ShowWindowDialogNotificationAllow='Allow';
Strings.OfficeOM.L_ShowWindowDialogNotificationIgnore='Ignore';
Strings.OfficeOM.L_ShuttingDown='Operation failed because the data is not current on the server.';
Strings.OfficeOM.L_SliceSizeNotSupported='The specified slice size is not supported.';
Strings.OfficeOM.L_SpecifiedIdNotExist='The specified ID does not exist.';
Strings.OfficeOM.L_Timeout='The operation has timed out.';
Strings.OfficeOM.L_TooManyArguments='too many arguments';
Strings.OfficeOM.L_TooManyIncompleteRequests='Wait until the previous call completes.';
Strings.OfficeOM.L_TooManyOptionalFunction='multiple optional functions in parameter list';
Strings.OfficeOM.L_TooManyOptionalObjects='multiple optional objects in parameter list';
Strings.OfficeOM.L_UnknownBindingType='The binding type is not supported.';
Strings.OfficeOM.L_UnsupportedDataObject='The supplied data object type is not supported.';
Strings.OfficeOM.L_UnsupportedEnumeration='Unsupported Enumeration';
Strings.OfficeOM.L_UnsupportedEnumerationMessage='The enumeration isn\'t supported in the current host application.';
Strings.OfficeOM.L_UnsupportedUserIdentity='User identity type is not supported.';
Strings.OfficeOM.L_UnsupportedUserIdentityMessage='The identity type of the user is not supported.';
Strings.OfficeOM.L_UserAborted='User aborted the consent request.';
Strings.OfficeOM.L_UserAbortedMessage='The user did not consent the add-in permissions.';
Strings.OfficeOM.L_UserClickIgnore='The user chose to ignore the dialog box.';
Strings.OfficeOM.L_UserNotSignedIn='No user is signed into Office.';
Strings.OfficeOM.L_ValueNotLoaded='The value of the result object has not been loaded yet. Before reading the value property, call \"context.sync()\" on the associated request context.';
Strings.OfficeOM.L_WorkbookHiddenMessage='The JavaScript API request failed because the workbook was hidden. Please unhide the workbook and try again.';


OSF = typeof OSF === "object" ? OSF : {};
OSF.DDA = OSF.DDA || {};

OSF.DDA.ErrorCodeManager = function()
{
    var _errorMappings = {};
    return{
            getErrorArgs: function OSF_DDA_ErrorCodeManager$getErrorArgs(errorCode)
            {
                var errorArgs = _errorMappings[errorCode];
                if(!errorArgs)
                    errorArgs = _errorMappings[this.errorCodes.ooeInternalError];
                else
                {
                    if(!errorArgs.name)
                        errorArgs.name = _errorMappings[this.errorCodes.ooeInternalError].name;
                    if(!errorArgs.message)
                        errorArgs.message = _errorMappings[this.errorCodes.ooeInternalError].message
                }
                return errorArgs
            },
            addErrorMessage: function OSF_DDA_ErrorCodeManager$addErrorMessage(errorCode, errorNameMessage)
            {
                _errorMappings[errorCode] = errorNameMessage
            },
            errorCodes: {
                ooeSuccess: 0,
                ooeChunkResult: 1,
                ooeCoercionTypeNotSupported: 1e3,
                ooeGetSelectionNotMatchDataType: 1001,
                ooeCoercionTypeNotMatchBinding: 1002,
                ooeInvalidGetRowColumnCounts: 1003,
                ooeSelectionNotSupportCoercionType: 1004,
                ooeInvalidGetStartRowColumn: 1005,
                ooeNonUniformPartialGetNotSupported: 1006,
                ooeGetDataIsTooLarge: 1008,
                ooeFileTypeNotSupported: 1009,
                ooeGetDataParametersConflict: 1010,
                ooeInvalidGetColumns: 1011,
                ooeInvalidGetRows: 1012,
                ooeInvalidReadForBlankRow: 1013,
                ooeUnsupportedDataObject: 2e3,
                ooeCannotWriteToSelection: 2001,
                ooeDataNotMatchSelection: 2002,
                ooeOverwriteWorksheetData: 2003,
                ooeDataNotMatchBindingSize: 2004,
                ooeInvalidSetStartRowColumn: 2005,
                ooeInvalidDataFormat: 2006,
                ooeDataNotMatchCoercionType: 2007,
                ooeDataNotMatchBindingType: 2008,
                ooeSetDataIsTooLarge: 2009,
                ooeNonUniformPartialSetNotSupported: 2010,
                ooeInvalidSetColumns: 2011,
                ooeInvalidSetRows: 2012,
                ooeSetDataParametersConflict: 2013,
                ooeCellDataAmountBeyondLimits: 2014,
                ooeSelectionCannotBound: 3e3,
                ooeBindingNotExist: 3002,
                ooeBindingToMultipleSelection: 3003,
                ooeInvalidSelectionForBindingType: 3004,
                ooeOperationNotSupportedOnThisBindingType: 3005,
                ooeNamedItemNotFound: 3006,
                ooeMultipleNamedItemFound: 3007,
                ooeInvalidNamedItemForBindingType: 3008,
                ooeUnknownBindingType: 3009,
                ooeOperationNotSupportedOnMatrixData: 3010,
                ooeInvalidColumnsForBinding: 3011,
                ooeSettingNameNotExist: 4e3,
                ooeSettingsCannotSave: 4001,
                ooeSettingsAreStale: 4002,
                ooeOperationNotSupported: 5e3,
                ooeInternalError: 5001,
                ooeDocumentReadOnly: 5002,
                ooeEventHandlerNotExist: 5003,
                ooeInvalidApiCallInContext: 5004,
                ooeShuttingDown: 5005,
                ooeUnsupportedEnumeration: 5007,
                ooeIndexOutOfRange: 5008,
                ooeBrowserAPINotSupported: 5009,
                ooeInvalidParam: 5010,
                ooeRequestTimeout: 5011,
                ooeInvalidOrTimedOutSession: 5012,
                ooeInvalidApiArguments: 5013,
                ooeOperationCancelled: 5014,
                ooeWorkbookHidden: 5015,
                ooeTooManyIncompleteRequests: 5100,
                ooeRequestTokenUnavailable: 5101,
                ooeActivityLimitReached: 5102,
                ooeCustomXmlNodeNotFound: 6e3,
                ooeCustomXmlError: 6100,
                ooeCustomXmlExceedQuota: 6101,
                ooeCustomXmlOutOfDate: 6102,
                ooeNoCapability: 7e3,
                ooeCannotNavTo: 7001,
                ooeSpecifiedIdNotExist: 7002,
                ooeNavOutOfBound: 7004,
                ooeElementMissing: 8e3,
                ooeProtectedError: 8001,
                ooeInvalidCellsValue: 8010,
                ooeInvalidTableOptionValue: 8011,
                ooeInvalidFormatValue: 8012,
                ooeRowIndexOutOfRange: 8020,
                ooeColIndexOutOfRange: 8021,
                ooeFormatValueOutOfRange: 8022,
                ooeCellFormatAmountBeyondLimits: 8023,
                ooeMemoryFileLimit: 11e3,
                ooeNetworkProblemRetrieveFile: 11001,
                ooeInvalidSliceSize: 11002,
                ooeInvalidCallback: 11101,
                ooeInvalidWidth: 12e3,
                ooeInvalidHeight: 12001,
                ooeNavigationError: 12002,
                ooeInvalidScheme: 12003,
                ooeAppDomains: 12004,
                ooeRequireHTTPS: 12005,
                ooeWebDialogClosed: 12006,
                ooeDialogAlreadyOpened: 12007,
                ooeEndUserAllow: 12008,
                ooeEndUserIgnore: 12009,
                ooeNotUILessDialog: 12010,
                ooeCrossZone: 12011,
                ooeNotSSOAgave: 13e3,
                ooeSSOUserNotSignedIn: 13001,
                ooeSSOUserAborted: 13002,
                ooeSSOUnsupportedUserIdentity: 13003,
                ooeSSOInvalidResourceUrl: 13004,
                ooeSSOInvalidGrant: 13005,
                ooeSSOClientError: 13006,
                ooeSSOServerError: 13007,
                ooeAddinIsAlreadyRequestingToken: 13008,
                ooeSSOUserConsentNotSupportedByCurrentAddinCategory: 13009,
                ooeSSOConnectionLost: 13010,
                ooeResourceNotAllowed: 13011,
                ooeSSOUnsupportedPlatform: 13012,
                ooeAccessDenied: 13990,
                ooeGeneralException: 13991
            },
            initializeErrorMessages: function OSF_DDA_ErrorCodeManager$initializeErrorMessages(stringNS)
            {
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCoercionTypeNotSupported] = {
                    name: stringNS.L_InvalidCoercion,
                    message: stringNS.L_CoercionTypeNotSupported
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeGetSelectionNotMatchDataType] = {
                    name: stringNS.L_DataReadError,
                    message: stringNS.L_GetSelectionNotSupported
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCoercionTypeNotMatchBinding] = {
                    name: stringNS.L_InvalidCoercion,
                    message: stringNS.L_CoercionTypeNotMatchBinding
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidGetRowColumnCounts] = {
                    name: stringNS.L_DataReadError,
                    message: stringNS.L_InvalidGetRowColumnCounts
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSelectionNotSupportCoercionType] = {
                    name: stringNS.L_DataReadError,
                    message: stringNS.L_SelectionNotSupportCoercionType
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidGetStartRowColumn] = {
                    name: stringNS.L_DataReadError,
                    message: stringNS.L_InvalidGetStartRowColumn
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNonUniformPartialGetNotSupported] = {
                    name: stringNS.L_DataReadError,
                    message: stringNS.L_NonUniformPartialGetNotSupported
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeGetDataIsTooLarge] = {
                    name: stringNS.L_DataReadError,
                    message: stringNS.L_GetDataIsTooLarge
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeFileTypeNotSupported] = {
                    name: stringNS.L_DataReadError,
                    message: stringNS.L_FileTypeNotSupported
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeGetDataParametersConflict] = {
                    name: stringNS.L_DataReadError,
                    message: stringNS.L_GetDataParametersConflict
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidGetColumns] = {
                    name: stringNS.L_DataReadError,
                    message: stringNS.L_InvalidGetColumns
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidGetRows] = {
                    name: stringNS.L_DataReadError,
                    message: stringNS.L_InvalidGetRows
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidReadForBlankRow] = {
                    name: stringNS.L_DataReadError,
                    message: stringNS.L_InvalidReadForBlankRow
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeUnsupportedDataObject] = {
                    name: stringNS.L_DataWriteError,
                    message: stringNS.L_UnsupportedDataObject
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCannotWriteToSelection] = {
                    name: stringNS.L_DataWriteError,
                    message: stringNS.L_CannotWriteToSelection
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeDataNotMatchSelection] = {
                    name: stringNS.L_DataWriteError,
                    message: stringNS.L_DataNotMatchSelection
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeOverwriteWorksheetData] = {
                    name: stringNS.L_DataWriteError,
                    message: stringNS.L_OverwriteWorksheetData
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeDataNotMatchBindingSize] = {
                    name: stringNS.L_DataWriteError,
                    message: stringNS.L_DataNotMatchBindingSize
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSetStartRowColumn] = {
                    name: stringNS.L_DataWriteError,
                    message: stringNS.L_InvalidSetStartRowColumn
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidDataFormat] = {
                    name: stringNS.L_InvalidFormat,
                    message: stringNS.L_InvalidDataFormat
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeDataNotMatchCoercionType] = {
                    name: stringNS.L_InvalidDataObject,
                    message: stringNS.L_DataNotMatchCoercionType
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeDataNotMatchBindingType] = {
                    name: stringNS.L_InvalidDataObject,
                    message: stringNS.L_DataNotMatchBindingType
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSetDataIsTooLarge] = {
                    name: stringNS.L_DataWriteError,
                    message: stringNS.L_SetDataIsTooLarge
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNonUniformPartialSetNotSupported] = {
                    name: stringNS.L_DataWriteError,
                    message: stringNS.L_NonUniformPartialSetNotSupported
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSetColumns] = {
                    name: stringNS.L_DataWriteError,
                    message: stringNS.L_InvalidSetColumns
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSetRows] = {
                    name: stringNS.L_DataWriteError,
                    message: stringNS.L_InvalidSetRows
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSetDataParametersConflict] = {
                    name: stringNS.L_DataWriteError,
                    message: stringNS.L_SetDataParametersConflict
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSelectionCannotBound] = {
                    name: stringNS.L_BindingCreationError,
                    message: stringNS.L_SelectionCannotBound
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeBindingNotExist] = {
                    name: stringNS.L_InvalidBindingError,
                    message: stringNS.L_BindingNotExist
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeBindingToMultipleSelection] = {
                    name: stringNS.L_BindingCreationError,
                    message: stringNS.L_BindingToMultipleSelection
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSelectionForBindingType] = {
                    name: stringNS.L_BindingCreationError,
                    message: stringNS.L_InvalidSelectionForBindingType
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeOperationNotSupportedOnThisBindingType] = {
                    name: stringNS.L_InvalidBindingOperation,
                    message: stringNS.L_OperationNotSupportedOnThisBindingType
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNamedItemNotFound] = {
                    name: stringNS.L_BindingCreationError,
                    message: stringNS.L_NamedItemNotFound
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeMultipleNamedItemFound] = {
                    name: stringNS.L_BindingCreationError,
                    message: stringNS.L_MultipleNamedItemFound
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidNamedItemForBindingType] = {
                    name: stringNS.L_BindingCreationError,
                    message: stringNS.L_InvalidNamedItemForBindingType
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeUnknownBindingType] = {
                    name: stringNS.L_InvalidBinding,
                    message: stringNS.L_UnknownBindingType
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeOperationNotSupportedOnMatrixData] = {
                    name: stringNS.L_InvalidBindingOperation,
                    message: stringNS.L_OperationNotSupportedOnMatrixData
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidColumnsForBinding] = {
                    name: stringNS.L_InvalidBinding,
                    message: stringNS.L_InvalidColumnsForBinding
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSettingNameNotExist] = {
                    name: stringNS.L_ReadSettingsError,
                    message: stringNS.L_SettingNameNotExist
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSettingsCannotSave] = {
                    name: stringNS.L_SaveSettingsError,
                    message: stringNS.L_SettingsCannotSave
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSettingsAreStale] = {
                    name: stringNS.L_SettingsStaleError,
                    message: stringNS.L_SettingsAreStale
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeOperationNotSupported] = {
                    name: stringNS.L_HostError,
                    message: stringNS.L_OperationNotSupported
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError] = {
                    name: stringNS.L_InternalError,
                    message: stringNS.L_InternalErrorDescription
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeDocumentReadOnly] = {
                    name: stringNS.L_PermissionDenied,
                    message: stringNS.L_DocumentReadOnly
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeEventHandlerNotExist] = {
                    name: stringNS.L_EventRegistrationError,
                    message: stringNS.L_EventHandlerNotExist
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidApiCallInContext] = {
                    name: stringNS.L_InvalidAPICall,
                    message: stringNS.L_InvalidApiCallInContext
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeShuttingDown] = {
                    name: stringNS.L_ShuttingDown,
                    message: stringNS.L_ShuttingDown
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeUnsupportedEnumeration] = {
                    name: stringNS.L_UnsupportedEnumeration,
                    message: stringNS.L_UnsupportedEnumerationMessage
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeIndexOutOfRange] = {
                    name: stringNS.L_IndexOutOfRange,
                    message: stringNS.L_IndexOutOfRange
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeBrowserAPINotSupported] = {
                    name: stringNS.L_APINotSupported,
                    message: stringNS.L_BrowserAPINotSupported
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeRequestTimeout] = {
                    name: stringNS.L_APICallFailed,
                    message: stringNS.L_RequestTimeout
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidOrTimedOutSession] = {
                    name: stringNS.L_InvalidOrTimedOutSession,
                    message: stringNS.L_InvalidOrTimedOutSessionMessage
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeTooManyIncompleteRequests] = {
                    name: stringNS.L_APICallFailed,
                    message: stringNS.L_TooManyIncompleteRequests
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeRequestTokenUnavailable] = {
                    name: stringNS.L_APICallFailed,
                    message: stringNS.L_RequestTokenUnavailable
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeActivityLimitReached] = {
                    name: stringNS.L_APICallFailed,
                    message: stringNS.L_ActivityLimitReached
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidApiArguments] = {
                    name: stringNS.L_APICallFailed,
                    message: stringNS.L_InvalidApiArgumentsMessage
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeWorkbookHidden] = {
                    name: stringNS.L_APICallFailed,
                    message: stringNS.L_WorkbookHiddenMessage
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCustomXmlNodeNotFound] = {
                    name: stringNS.L_InvalidNode,
                    message: stringNS.L_CustomXmlNodeNotFound
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCustomXmlError] = {
                    name: stringNS.L_CustomXmlError,
                    message: stringNS.L_CustomXmlError
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCustomXmlExceedQuota] = {
                    name: stringNS.L_CustomXmlExceedQuotaName,
                    message: stringNS.L_CustomXmlExceedQuotaMessage
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCustomXmlOutOfDate] = {
                    name: stringNS.L_CustomXmlOutOfDateName,
                    message: stringNS.L_CustomXmlOutOfDateMessage
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNoCapability] = {
                    name: stringNS.L_PermissionDenied,
                    message: stringNS.L_NoCapability
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCannotNavTo] = {
                    name: stringNS.L_CannotNavigateTo,
                    message: stringNS.L_CannotNavigateTo
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSpecifiedIdNotExist] = {
                    name: stringNS.L_SpecifiedIdNotExist,
                    message: stringNS.L_SpecifiedIdNotExist
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNavOutOfBound] = {
                    name: stringNS.L_NavOutOfBound,
                    message: stringNS.L_NavOutOfBound
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCellDataAmountBeyondLimits] = {
                    name: stringNS.L_DataWriteReminder,
                    message: stringNS.L_CellDataAmountBeyondLimits
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeElementMissing] = {
                    name: stringNS.L_MissingParameter,
                    message: stringNS.L_ElementMissing
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeProtectedError] = {
                    name: stringNS.L_PermissionDenied,
                    message: stringNS.L_NoCapability
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidCellsValue] = {
                    name: stringNS.L_InvalidValue,
                    message: stringNS.L_InvalidCellsValue
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidTableOptionValue] = {
                    name: stringNS.L_InvalidValue,
                    message: stringNS.L_InvalidTableOptionValue
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidFormatValue] = {
                    name: stringNS.L_InvalidValue,
                    message: stringNS.L_InvalidFormatValue
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeRowIndexOutOfRange] = {
                    name: stringNS.L_OutOfRange,
                    message: stringNS.L_RowIndexOutOfRange
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeColIndexOutOfRange] = {
                    name: stringNS.L_OutOfRange,
                    message: stringNS.L_ColIndexOutOfRange
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeFormatValueOutOfRange] = {
                    name: stringNS.L_OutOfRange,
                    message: stringNS.L_FormatValueOutOfRange
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCellFormatAmountBeyondLimits] = {
                    name: stringNS.L_FormattingReminder,
                    message: stringNS.L_CellFormatAmountBeyondLimits
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeMemoryFileLimit] = {
                    name: stringNS.L_MemoryLimit,
                    message: stringNS.L_CloseFileBeforeRetrieve
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNetworkProblemRetrieveFile] = {
                    name: stringNS.L_NetworkProblem,
                    message: stringNS.L_NetworkProblemRetrieveFile
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidSliceSize] = {
                    name: stringNS.L_InvalidValue,
                    message: stringNS.L_SliceSizeNotSupported
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeDialogAlreadyOpened] = {
                    name: stringNS.L_DisplayDialogError,
                    message: stringNS.L_DialogAlreadyOpened
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidWidth] = {
                    name: stringNS.L_IndexOutOfRange,
                    message: stringNS.L_IndexOutOfRange
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidHeight] = {
                    name: stringNS.L_IndexOutOfRange,
                    message: stringNS.L_IndexOutOfRange
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNavigationError] = {
                    name: stringNS.L_DisplayDialogError,
                    message: stringNS.L_NetworkProblem
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeInvalidScheme] = {
                    name: stringNS.L_DialogNavigateError,
                    message: stringNS.L_DialogInvalidScheme
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeAppDomains] = {
                    name: stringNS.L_DisplayDialogError,
                    message: stringNS.L_DialogAddressNotTrusted
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeRequireHTTPS] = {
                    name: stringNS.L_DisplayDialogError,
                    message: stringNS.L_DialogRequireHTTPS
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeEndUserIgnore] = {
                    name: stringNS.L_DisplayDialogError,
                    message: stringNS.L_UserClickIgnore
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeCrossZone] = {
                    name: stringNS.L_DisplayDialogError,
                    message: stringNS.L_NewWindowCrossZoneErrorString
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeNotSSOAgave] = {
                    name: stringNS.L_APINotSupported,
                    message: stringNS.L_InvalidSSOAddinMessage
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOUserNotSignedIn] = {
                    name: stringNS.L_UserNotSignedIn,
                    message: stringNS.L_UserNotSignedIn
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOUserAborted] = {
                    name: stringNS.L_UserAborted,
                    message: stringNS.L_UserAbortedMessage
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOUnsupportedUserIdentity] = {
                    name: stringNS.L_UnsupportedUserIdentity,
                    message: stringNS.L_UnsupportedUserIdentityMessage
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOInvalidResourceUrl] = {
                    name: stringNS.L_InvalidResourceUrl,
                    message: stringNS.L_InvalidResourceUrlMessage
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOInvalidGrant] = {
                    name: stringNS.L_InvalidGrant,
                    message: stringNS.L_InvalidGrantMessage
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOClientError] = {
                    name: stringNS.L_SSOClientError,
                    message: stringNS.L_SSOClientErrorMessage
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOServerError] = {
                    name: stringNS.L_SSOServerError,
                    message: stringNS.L_SSOServerErrorMessage
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeAddinIsAlreadyRequestingToken] = {
                    name: stringNS.L_AddinIsAlreadyRequestingToken,
                    message: stringNS.L_AddinIsAlreadyRequestingTokenMessage
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOUserConsentNotSupportedByCurrentAddinCategory] = {
                    name: stringNS.L_SSOUserConsentNotSupportedByCurrentAddinCategory,
                    message: stringNS.L_SSOUserConsentNotSupportedByCurrentAddinCategoryMessage
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOConnectionLost] = {
                    name: stringNS.L_SSOConnectionLostError,
                    message: stringNS.L_SSOConnectionLostErrorMessage
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeSSOUnsupportedPlatform] = {
                    name: stringNS.L_APINotSupported,
                    message: stringNS.L_SSOUnsupportedPlatform
                };
                _errorMappings[OSF.DDA.ErrorCodeManager.errorCodes.ooeOperationCancelled] = {
                    name: stringNS.L_OperationCancelledError,
                    message: stringNS.L_OperationCancelledErrorMessage
                }
            }
        }
}();


_u = typeof _u === "object" ? _u : {};
_u.ExtensibilityStrings = _u.ExtensibilityStrings || {};

_u.ExtensibilityStrings.l_APICallFailedDueToItemChange_Text='The selected item has been changed.';
_u.ExtensibilityStrings.l_ActionsDefinitionMultipleActionsError_Text='Only a single action is currently supported.';
_u.ExtensibilityStrings.l_ActionsDefinitionWrongNotificationMessageError_Text='Actions can\'t be defined for this notification message type.';
_u.ExtensibilityStrings.l_AttachmentDeleteGeneralFailure_Text='The attachment cannot be deleted from the item.';
_u.ExtensibilityStrings.l_AttachmentDeletedBeforeUploadCompletes_Text='The user removed the attachment before upload has completed.';
_u.ExtensibilityStrings.l_AttachmentErrorName_Text='Attachment Error';
_u.ExtensibilityStrings.l_AttachmentExceededSize_Text='The attachment cannot be added because it is too big.';
_u.ExtensibilityStrings.l_AttachmentItemIdTooLong_Text='One or more of the attachment IDs is too long.';
_u.ExtensibilityStrings.l_AttachmentNameTooLong_Text='One or more of the attachment names is too long.';
_u.ExtensibilityStrings.l_AttachmentNotSupported_Text='The attachment type is not supported.';
_u.ExtensibilityStrings.l_AttachmentUploadGeneralFailure_Text='The attachment cannot be added to the item.';
_u.ExtensibilityStrings.l_AttachmentUrlTooLong_Text='One or more of the attachment URLs is too long.';
_u.ExtensibilityStrings.l_CallSaveAsyncBeforeToken_Text='The token can\'t be retrieved until the item is saved.';
_u.ExtensibilityStrings.l_CannotAddAttachmentBeforeUpgrade_Text='Attachments cannot be added while the full reply or forward is being retrieved from the server.';
_u.ExtensibilityStrings.l_CannotPersistPropertyInUnsavedDraftError_Text='Notifications can\'t be persisted in unsaved drafts. Save the item before you call this API.';
_u.ExtensibilityStrings.l_CursorPositionChanged_Text='The user changed the position of the cursor while the data was being inserted.';
_u.ExtensibilityStrings.l_DataWriteErrorName_Text='Data Write Error';
_u.ExtensibilityStrings.l_DeleteAttachmentDoesNotExist_Text='The attachment cannot be deleted because attachment with the attachment index cannot be found.';
_u.ExtensibilityStrings.l_DisplayNameTooLong_Text='One or more of the display names provided is too long.';
_u.ExtensibilityStrings.l_DisplayReplyFormHtmlBodyRequired_Text='The \'htmlBody\' is required.';
_u.ExtensibilityStrings.l_DuplicateNotificationKey_Text='A notification with the provided key already exists.';
_u.ExtensibilityStrings.l_Duplicate_Category_Error_Text='One of the categories provided is already in the master category list.';
_u.ExtensibilityStrings.l_ElevatedPermissionNeededForMethod_Text='Elevated permission is required to call the method: \'{0}\'.';
_u.ExtensibilityStrings.l_ElevatedPermissionNeeded_Text='Elevated permission is required to access protected members of the JavaScript API for Office.';
_u.ExtensibilityStrings.l_EmailAddressTooLong_Text='One or more of the email addresses provided is too long.';
_u.ExtensibilityStrings.l_EwsRequestOversized_Text='The request exceeds the 1 MB size limit. Please modify your EWS request.';
_u.ExtensibilityStrings.l_ExceededMaxNumberOfAttachments_Text='Attachments cannot be added because the message already has the maximum number of attachments';
_u.ExtensibilityStrings.l_HtmlSanitizationFailure_Text='The HTML santization has failed.';
_u.ExtensibilityStrings.l_Insufficient_Item_Permissions_Text='The user doesn\'t have the permissions required to perform this operation.';
_u.ExtensibilityStrings.l_InternalFormatError_Text='There was an internal format error.';
_u.ExtensibilityStrings.l_InternalProtocolError_Text='Internal protocol error: \'{0}\'.';
_u.ExtensibilityStrings.l_Internal_Server_Error_Text='The Exchange server returned an error. Please look at the diagnostics object for more information.';
_u.ExtensibilityStrings.l_Internet_Not_Connected_Error_Text='The user is no longer connected to the network. Please check your network connection and try again.';
_u.ExtensibilityStrings.l_InvalidActionType_Text='The value of the parameter \'actionType\' is invalid. The accepted value is \"showTaskPane\".';
_u.ExtensibilityStrings.l_InvalidAttachmentId_Text='The attachment ID was invalid.';
_u.ExtensibilityStrings.l_InvalidAttachmentPath_Text='The attachment path was invalid.';
_u.ExtensibilityStrings.l_InvalidCommandIdError_Text='The value of the parameter \'commandId\' is invalid.';
_u.ExtensibilityStrings.l_InvalidDate_Text='The input doesn\'t resolve to a valid date.';
_u.ExtensibilityStrings.l_InvalidEndTime_Text='The end time cannot be before the start time.';
_u.ExtensibilityStrings.l_InvalidEventDates_Text='The end date occurs before the start date.';
_u.ExtensibilityStrings.l_InvalidParameterValueError_Text='The value of the \'{0}\' parameter is invalid.';
_u.ExtensibilityStrings.l_InvalidSelection_Text='The selection is invalid.';
_u.ExtensibilityStrings.l_InvalidTime_Text='The input doesn\'t resolve to a valid time.';
_u.ExtensibilityStrings.l_Invalid_Category_Error_Text='Invalid categories were provided.';
_u.ExtensibilityStrings.l_Item_Not_Saved_Error_Text='The id can\'t be retrieved until the item is saved.';
_u.ExtensibilityStrings.l_MessageInDifferentStoreError_Text='The EWS ID can\'t be retrieved because the message is saved in a different store.';
_u.ExtensibilityStrings.l_NoValidRecipientsProvided_Text='No valid recipients were provided.';
_u.ExtensibilityStrings.l_NotificationKeyNotFound_Text='There are no notifications with the provided key.';
_u.ExtensibilityStrings.l_NullOrEmptyParameterError_Text='The parameter \'{0}\' is mandatory and it can\'t be null or empty.';
_u.ExtensibilityStrings.l_NumberOfNotificationsExceeded_Text='The notification couldn\'t be added because the notification limit has been reached.';
_u.ExtensibilityStrings.l_NumberOfRecipientsExceeded_Text='The total number of recipients in the field can\'t exceed {0}.';
_u.ExtensibilityStrings.l_OffsetNotfound_Text='An offset for this time stamp couldn\'t be found.';
_u.ExtensibilityStrings.l_Olk_Http_Error_Text='The request has failed. Please look at the diagnostics object for the HTTP error code.';
_u.ExtensibilityStrings.l_ParameterValueTooLongError_Text='The value of the \'{0}\' parameter is too long. The maximum number of characters is \'{1}\'.';
_u.ExtensibilityStrings.l_ParametersNotAsExpected_Text='The given parameters do not match the expected format.';
_u.ExtensibilityStrings.l_PersistedNotificationArrayReadError_Text='The API call you made failed because the persisted notifications couldn\'t be retrieved.';
_u.ExtensibilityStrings.l_PersistedNotificationArraySaveError_Text='The API call you made failed because notifications couldn\'t be persisted.';
_u.ExtensibilityStrings.l_RecurrenceErrorMaxOccurrences_Text='The recurring series exceeds the maximum limit of 999 occurrences.';
_u.ExtensibilityStrings.l_RecurrenceErrorZeroOccurrences_Text='The recurring series has no occurrences in the specified date range.';
_u.ExtensibilityStrings.l_RecurrenceInvalidTimeZone_Text='The specified time zone is not supported.';
_u.ExtensibilityStrings.l_RecurrenceUnsupportedAlternateCalendar_Text='The recurrence pattern was set by the user using an alternate calendar that is not supported.';
_u.ExtensibilityStrings.l_Recurrence_Error_Instance_SetAsync_Text='A recurrence pattern can\'t be set for a single occurrence in a series.';
_u.ExtensibilityStrings.l_Recurrence_Error_Properties_Invalid_Text='The recurrence pattern isn\'t valid. Please check that the specified recurrence properties align with the recurrence type.';
_u.ExtensibilityStrings.l_SaveError_Text='Connection error occurred while trying to save the item on the server.';


Microsoft = typeof Microsoft === "object" ? Microsoft : {}
Microsoft.Office = Microsoft.Office || {};
Microsoft.Office.WebExtension = Microsoft.Office.WebExtension || {};
Microsoft.Office.WebExtension.MailboxEnums = Microsoft.Office.WebExtension.MailboxEnums || {};

OSF = typeof OSF === "object" ? OSF : {}
OSF.InitializationHelper = OSF.InitializationHelper || {};

Office = typeof Office === "object" ? Office : {};
Office.context = Office.context || {};
Office.context.roamingSettings = Office.context.roamingSettings || {};

var MailboxCreateOMHelper = function(appContext, appReady, basePath) {

var Outlook = typeof Outlook === "object" ? Outlook : {}; Outlook["OutlookAppOm"] =
/******/ (function(modules) { // webpackBootstrap
/******/ 	// The module cache
/******/ 	var installedModules = {};
/******/
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/
/******/ 		// Check if module is in cache
/******/ 		if(installedModules[moduleId]) {
/******/ 			return installedModules[moduleId].exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = installedModules[moduleId] = {
/******/ 			i: moduleId,
/******/ 			l: false,
/******/ 			exports: {}
/******/ 		};
/******/
/******/ 		// Execute the module function
/******/ 		modules[moduleId].call(module.exports, module, module.exports, __webpack_require__);
/******/
/******/ 		// Flag the module as loaded
/******/ 		module.l = true;
/******/
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/
/******/
/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = modules;
/******/
/******/ 	// expose the module cache
/******/ 	__webpack_require__.c = installedModules;
/******/
/******/ 	// define getter function for harmony exports
/******/ 	__webpack_require__.d = function(exports, name, getter) {
/******/ 		if(!__webpack_require__.o(exports, name)) {
/******/ 			Object.defineProperty(exports, name, { enumerable: true, get: getter });
/******/ 		}
/******/ 	};
/******/
/******/ 	// define __esModule on exports
/******/ 	__webpack_require__.r = function(exports) {
/******/ 		if(typeof Symbol !== 'undefined' && Symbol.toStringTag) {
/******/ 			Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/ 		}
/******/ 		Object.defineProperty(exports, '__esModule', { value: true });
/******/ 	};
/******/
/******/ 	// create a fake namespace object
/******/ 	// mode & 1: value is a module id, require it
/******/ 	// mode & 2: merge all properties of value into the ns
/******/ 	// mode & 4: return value when already ns object
/******/ 	// mode & 8|1: behave like require
/******/ 	__webpack_require__.t = function(value, mode) {
/******/ 		if(mode & 1) value = __webpack_require__(value);
/******/ 		if(mode & 8) return value;
/******/ 		if((mode & 4) && typeof value === 'object' && value && value.__esModule) return value;
/******/ 		var ns = Object.create(null);
/******/ 		__webpack_require__.r(ns);
/******/ 		Object.defineProperty(ns, 'default', { enumerable: true, value: value });
/******/ 		if(mode & 2 && typeof value != 'string') for(var key in value) __webpack_require__.d(ns, key, function(key) { return value[key]; }.bind(null, key));
/******/ 		return ns;
/******/ 	};
/******/
/******/ 	// getDefaultExport function for compatibility with non-harmony modules
/******/ 	__webpack_require__.n = function(module) {
/******/ 		var getter = module && module.__esModule ?
/******/ 			function getDefault() { return module['default']; } :
/******/ 			function getModuleExports() { return module; };
/******/ 		__webpack_require__.d(getter, 'a', getter);
/******/ 		return getter;
/******/ 	};
/******/
/******/ 	// Object.prototype.hasOwnProperty.call
/******/ 	__webpack_require__.o = function(object, property) { return Object.prototype.hasOwnProperty.call(object, property); };
/******/
/******/ 	// __webpack_public_path__
/******/ 	__webpack_require__.p = "/";
/******/
/******/
/******/ 	// Load entry module and return exports
/******/ 	return __webpack_require__(__webpack_require__.s = 1);
/******/ })
/************************************************************************/
/******/ ([
/* 0 */
/***/ (function(module, exports) {

module.exports = OSF;

/***/ }),
/* 1 */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);

// CONCATENATED MODULE: ./src/utils/getErrorForTelemetry.ts
var getErrorForTelemetry = function getErrorForTelemetry(resultCode, responseDictionary) {
  return resultCode;
};
// CONCATENATED MODULE: ./src/utils/isOwaOnly.ts
var isOwaOnly = function isOwaOnly(dispid) {
  switch (dispid) {
    case 402:
    case 401:
    case 400:
    case 403:
      return true;

    default:
      return false;
  }
};
// CONCATENATED MODULE: ./src/utils/InvokeResultCode.ts
var InvokeResultCode;

(function (InvokeResultCode) {
  InvokeResultCode[InvokeResultCode["noError"] = 0] = "noError";
  InvokeResultCode[InvokeResultCode["errorInRequest"] = -1] = "errorInRequest";
  InvokeResultCode[InvokeResultCode["errorHandlingRequest"] = -2] = "errorHandlingRequest";
  InvokeResultCode[InvokeResultCode["errorInResponse"] = -3] = "errorInResponse";
  InvokeResultCode[InvokeResultCode["errorHandlingResponse"] = -4] = "errorHandlingResponse";
  InvokeResultCode[InvokeResultCode["errorHandlingRequestAccessDenied"] = -5] = "errorHandlingRequestAccessDenied";
  InvokeResultCode[InvokeResultCode["errorHandlingMethodCallTimedout"] = -6] = "errorHandlingMethodCallTimedout";
})(InvokeResultCode || (InvokeResultCode = {}));
// CONCATENATED MODULE: ./src/utils/isNullOrUndefined.ts
function isNullOrUndefined(value) {
  return value === null || value === undefined;
}
// CONCATENATED MODULE: ./src/types/ExtensibilityStrings.ts

var OfficeStringJS = "office_strings.js";
var OfficeStringDebugJS = "office_strings.debug.js";
var ExtensibilityStringJS = "outlook_strings.js";
var tempWindow = window;
var ExtensibilityStrings;
function getString(string) {
  return ExtensibilityStrings[string];
}
var ExtensibilityStrings_url = "";
var baseUrl = "";
var scriptElement = null;
var stringLoadedCallback;
var stringsAreLoaded = false;

function createScriptElement(url) {
  var scriptElement = document.createElement("script");
  scriptElement.type = "text/javascript";
  scriptElement.src = url;
  return scriptElement;
}

function loadLocalizedScript(initializeAppCallback) {
  stringLoadedCallback = initializeAppCallback;
  var officeIndex;
  var scripts = document.getElementsByTagName("script");

  for (var i = 0; i < scripts.length; i++) {
    var tag = scripts.item(i);

    if (tag && tag.src) {
      var filename = tag.src || "";
      filename = filename.toLowerCase();
      officeIndex = filename.indexOf(OfficeStringJS);

      if (filename && officeIndex > 0) {
        ExtensibilityStrings_url = filename.replace(OfficeStringJS, ExtensibilityStringJS);
        baseUrl = saveBaseUrl(baseUrl, officeIndex, filename);
        break;
      }

      officeIndex = filename.indexOf(OfficeStringDebugJS);

      if (filename && officeIndex > 0) {
        ExtensibilityStrings_url = filename.replace(OfficeStringDebugJS, ExtensibilityStringJS);
        baseUrl = saveBaseUrl(baseUrl, officeIndex, filename);
        break;
      }
    }
  }

  if (ExtensibilityStrings_url) {
    var head = document.getElementsByTagName("head")[0];
    scriptElement = createScriptElement(ExtensibilityStrings_url);
    scriptElement.onload = scriptElementCallback;
    scriptElement.onreadystatechange = scriptElementCallback;
    window.setTimeout(failureCallback, 2000);
    head.appendChild(scriptElement);
  }
}

function scriptElementCallback() {
  stringsAreLoaded = true;

  if (!isNullOrUndefined(stringLoadedCallback) && (isNullOrUndefined(scriptElement.readyState) || !isNullOrUndefined(scriptElement.readyState) && (scriptElement.readyState === "loaded" || scriptElement.readyState === "complete"))) {
    scriptElement.onload = null;
    scriptElement.onreadystatechange = null;

    if (typeof tempWindow._u !== "undefined") {
      ExtensibilityStrings = tempWindow._u.ExtensibilityStrings;
    }

    stringLoadedCallback();
  }
}

function failureCallback() {
  if (!stringsAreLoaded) {
    var head = document.getElementsByTagName("head")[0];
    var fallbackUrl = baseUrl + "en-us/" + ExtensibilityStringJS;
    scriptElement.onload = null;
    scriptElement.onreadystatechange = null;
    scriptElement = createScriptElement(fallbackUrl);
    scriptElement.onload = scriptElementCallback;
    scriptElement.onreadystatechange = scriptElementCallback;
    head.appendChild(scriptElement);
  }
}

function saveBaseUrl(baseUrl, officeIndex, filename) {
  var languageUrl = filename.substring(0, officeIndex);
  var lastIndexOfSlash = languageUrl.lastIndexOf("/", languageUrl.length - 2);

  if (lastIndexOfSlash === -1) {
    lastIndexOfSlash = languageUrl.lastIndexOf("\\", languageUrl.length - 2);
  }

  if (lastIndexOfSlash !== -1 && languageUrl.length > lastIndexOfSlash + 1) {
    baseUrl = languageUrl.substring(0, lastIndexOfSlash + 1);
  }

  return baseUrl;
}
// CONCATENATED MODULE: ./src/utils/getErrorArgs.ts


var getErrorArgs_OSF = __webpack_require__(0);

var isInitialized = false;
function getErrorArgs(detailedErrorCode) {
  if (!isInitialized) {
    initialize();
  }

  return getErrorArgs_OSF.DDA.ErrorCodeManager.getErrorArgs(detailedErrorCode);
}
var totalRecipientsLimit = 500;
function initialize() {
  addErrorMessage(9000, "AttachmentSizeExceeded", getString("l_AttachmentExceededSize_Text"));
  addErrorMessage(9001, "NumberOfAttachmentsExceeded", getString("l_ExceededMaxNumberOfAttachments_Text"));
  addErrorMessage(9002, "InternalFormatError", getString("l_InternalFormatError_Text"));
  addErrorMessage(9003, "InvalidAttachmentId", getString("l_InvalidAttachmentId_Text"));
  addErrorMessage(9004, "InvalidAttachmentPath", getString("l_InvalidAttachmentPath_Text"));
  addErrorMessage(9005, "CannotAddAttachmentBeforeUpgrade", getString("l_CannotAddAttachmentBeforeUpgrade_Text"));
  addErrorMessage(9006, "AttachmentDeletedBeforeUploadCompletes", getString("l_AttachmentDeletedBeforeUploadCompletes_Text"));
  addErrorMessage(9007, "AttachmentUploadGeneralFailure", getString("l_AttachmentUploadGeneralFailure_Text"));
  addErrorMessage(9008, "AttachmentToDeleteDoesNotExist", getString("l_DeleteAttachmentDoesNotExist_Text"));
  addErrorMessage(9009, "AttachmentDeleteGeneralFailure", getString("l_AttachmentDeleteGeneralFailure_Text"));
  addErrorMessage(9010, "InvalidEndTime", getString("l_InvalidEndTime_Text"));
  addErrorMessage(9011, "HtmlSanitizationFailure", getString("l_HtmlSanitizationFailure_Text"));
  addErrorMessage(9012, "NumberOfRecipientsExceeded", getString("l_NumberOfRecipientsExceeded_Text").replace("{0}", totalRecipientsLimit));
  addErrorMessage(9013, "NoValidRecipientsProvided", getString("l_NoValidRecipientsProvided_Text"));
  addErrorMessage(9014, "CursorPositionChanged", getString("l_CursorPositionChanged_Text"));
  addErrorMessage(9016, "InvalidSelection", getString("l_InvalidSelection_Text"));
  addErrorMessage(9017, "AccessRestricted", "");
  addErrorMessage(9018, "GenericTokenError", "");
  addErrorMessage(9019, "GenericSettingsError", "");
  addErrorMessage(9020, "GenericResponseError", "");
  addErrorMessage(9021, "SaveError", getString("l_SaveError_Text"));
  addErrorMessage(9022, "MessageInDifferentStoreError", getString("l_MessageInDifferentStoreError_Text"));
  addErrorMessage(9023, "DuplicateNotificationKey", getString("l_DuplicateNotificationKey_Text"));
  addErrorMessage(9024, "NotificationKeyNotFound", getString("l_NotificationKeyNotFound_Text"));
  addErrorMessage(9025, "NumberOfNotificationsExceeded", getString("l_NumberOfNotificationsExceeded_Text"));
  addErrorMessage(9026, "PersistedNotificationArrayReadError", getString("l_PersistedNotificationArrayReadError_Text"));
  addErrorMessage(9027, "PersistedNotificationArraySaveError", getString("l_PersistedNotificationArraySaveError_Text"));
  addErrorMessage(9028, "CannotPersistPropertyInUnsavedDraftError", getString("l_CannotPersistPropertyInUnsavedDraftError_Text"));
  addErrorMessage(9029, "CanOnlyGetTokenForSavedItem", getString("l_CallSaveAsyncBeforeToken_Text"));
  addErrorMessage(9030, "APICallFailedDueToItemChange", getString("l_APICallFailedDueToItemChange_Text"));
  addErrorMessage(9031, "InvalidParameterValueError", getString("l_InvalidParameterValueError_Text"));
  addErrorMessage(9032, "ApiCallNotSupportedByExtensionPoint", getString("l_API_Not_Supported_By_ExtensionPoint_Error_Text"));
  addErrorMessage(9033, "SetRecurrenceOnInstanceError", getString("l_Recurrence_Error_Instance_SetAsync_Text"));
  addErrorMessage(9034, "InvalidRecurrenceError", getString("l_Recurrence_Error_Properties_Invalid_Text"));
  addErrorMessage(9035, "RecurrenceZeroOccurrences", getString("l_RecurrenceErrorZeroOccurrences_Text"));
  addErrorMessage(9036, "RecurrenceMaxOccurrences", getString("l_RecurrenceErrorMaxOccurrences_Text"));
  addErrorMessage(9037, "RecurrenceInvalidTimeZone", getString("l_RecurrenceInvalidTimeZone_Text"));
  addErrorMessage(9038, "InsufficientItemPermissionsError", getString("l_Insufficient_Item_Permissions_Text"));
  addErrorMessage(9039, "RecurrenceUnsupportedAlternateCalendar", getString("l_RecurrenceUnsupportedAlternateCalendar_Text"));
  addErrorMessage(9040, "HTTPRequestFailure", getString("l_Olk_Http_Error_Text"));
  addErrorMessage(9041, "NetworkError", getString("l_Internet_Not_Connected_Error_Text"));
  addErrorMessage(9042, "InternalServerError", getString("l_Internal_Server_Error_Text"));
  addErrorMessage(9043, "AttachmentTypeNotSupported", getString("l_AttachmentNotSupported_Text"));
  addErrorMessage(9044, "InvalidCategory", getString("l_Invalid_Category_Error_Text"));
  addErrorMessage(9045, "DuplicateCategory", getString("l_Duplicate_Category_Error_Text"));
  addErrorMessage(9046, "ItemNotSaved", getString("l_Item_Not_Saved_Error_Text"));
  addErrorMessage(9047, "MissingExtendedPermissionsForAPIError", getString("l_Missing_Extended_Permissions_For_API"));
  addErrorMessage(9048, "TokenAccessDenied", getString("l_TokenAccessDeniedWithoutItemContext_Text"));
  addErrorMessage(9049, "ItemNotFound", getString("l_ItemNotFound_Text"));
  isInitialized = true;
}
function addErrorMessage(code, error, message) {
  getErrorArgs_OSF.DDA.ErrorCodeManager.addErrorMessage(code, {
    name: error,
    message: message
  });
}
// CONCATENATED MODULE: ./src/utils/AdditionalGlobalParameters.ts
var additionalOutlookGlobalParameters;
var getAdditionalGlobalParametersSingleton = function getAdditionalGlobalParametersSingleton() {
  return additionalOutlookGlobalParameters;
};
var recreateAdditionalGlobalParametersSingleton = function recreateAdditionalGlobalParametersSingleton(parameterBlobSupported) {
  additionalOutlookGlobalParameters = new AdditionalGlobalParameters();
  additionalOutlookGlobalParameters.parameterBlobSupported = true;
  return additionalOutlookGlobalParameters;
};

var AdditionalGlobalParameters = function () {
  function AdditionalGlobalParameters() {
    this._parameterBlobSupported = true;
    this._itemNumber = 0;
    additionalOutlookGlobalParameters = this;
  }

  Object.defineProperty(AdditionalGlobalParameters.prototype, "parameterBlobSupported", {
    set: function set(supported) {
      this._parameterBlobSupported = supported;
    },
    enumerable: true,
    configurable: true
  });

  AdditionalGlobalParameters.prototype.setActionsDefinition = function (actionsDefinitionIn) {
    this._actionsDefinition = actionsDefinitionIn;
  };

  AdditionalGlobalParameters.prototype.setCurrentItemNumber = function (itemNumberIn) {
    if (itemNumberIn > 0) {
      this._itemNumber = itemNumberIn;
    }
  };

  Object.defineProperty(AdditionalGlobalParameters.prototype, "itemNumber", {
    get: function get() {
      return this._itemNumber;
    },
    enumerable: true,
    configurable: true
  });
  Object.defineProperty(AdditionalGlobalParameters.prototype, "actionsDefinition", {
    get: function get() {
      return this._actionsDefinition;
    },
    enumerable: true,
    configurable: true
  });

  AdditionalGlobalParameters.prototype.updateOutlookExecuteParameters = function (executeParameters, additionalApiParameters) {
    var outParameters = executeParameters;

    if (this._parameterBlobSupported) {
      if (this._itemNumber > 0) {
        additionalApiParameters.itemNumber = this._itemNumber.toString();
      }

      if (this._actionsDefinition != null) {
        additionalApiParameters.actions = this.actionsDefinition;
      }

      if (Object.keys(additionalApiParameters).length === 0) {
        return outParameters;
      }

      if (outParameters == null) {
        outParameters = [];
      }

      outParameters.push(JSON.stringify(additionalApiParameters));
    }

    return outParameters;
  };

  return AdditionalGlobalParameters;
}();


// CONCATENATED MODULE: ./src/utils/callOutlookNativeDispatcher.ts


var callOutlookNativeDispatcher_OSF = __webpack_require__(0);

var callOutlookNativeDispatcher_callOutlookNativeDispatcher = function callOutlookNativeDispatcher(dispid, data, responseCallback) {
  var executeParameters = callOutlookNativeDispatcher_convertToOutlookNativeParameters(dispid, data);
  callOutlookNativeDispatcher_OSF.ClientHostController.execute(dispid, executeParameters, function (nativeData, resultCode) {
    var deserializedData = null;
    var responseData = nativeData.toArray();

    if (responseData.length > 0) {
      var itemNumberFromOutlookResponse = getItemNumberFromOutlookResponse(responseData);
      var isValidItemNumber = itemNumberFromOutlookResponse > 0;
      var itemChanged = isValidItemNumber && itemNumberFromOutlookResponse > getAdditionalGlobalParametersSingleton().itemNumber;
      deserializedData = createDeserializedData(responseData, itemChanged);
    }

    if (responseCallback != null) {
      responseCallback(resultCode, deserializedData);
    }
  });
};
var callOutlookNativeDispatcher_convertToOutlookNativeParameters = function convertToOutlookNativeParameters(dispid, data) {
  var executeParameters = null;
  var optionalParameters = {};

  switch (dispid) {
    case 12:
      optionalParameters.isRest = data.isRest;
      break;

    case 4:
      {
        var jsonProperty = JSON.stringify(data.customProperties);
        executeParameters = [jsonProperty];
        break;
      }

    case 5:
      executeParameters = new Array(data.body);
      break;

    case 8:
    case 9:
    case 179:
    case 180:
      executeParameters = new Array(data.itemId);
      break;

    case 7:
    case 177:
      executeParameters = new Array(convertRecipientArrayParameterForOutlookForDisplayApi(data.requiredAttendees), convertRecipientArrayParameterForOutlookForDisplayApi(data.optionalAttendees), data.start, data.end, data.location, convertRecipientArrayParameterForOutlookForDisplayApi(data.resources), data.subject, data.body);
      break;

    case 44:
    case 178:
      executeParameters = [convertRecipientArrayParameterForOutlookForDisplayApi(data.toRecipients), convertRecipientArrayParameterForOutlookForDisplayApi(data.ccRecipients), convertRecipientArrayParameterForOutlookForDisplayApi(data.bccRecipients), data.subject, data.htmlBody, data.attachments];
      break;

    case 43:
      executeParameters = [data.ewsIdOrEmail];
      break;

    case 45:
      executeParameters = [data.module, data.queryString];
      break;

    case 40:
      executeParameters = [data.extensionId, data.consentState];
      break;

    case 11:
    case 10:
    case 184:
    case 183:
      executeParameters = [data.htmlBody];
      break;

    case 31:
    case 30:
    case 182:
    case 181:
      executeParameters = [data.htmlBody, data.attachments];
      break;

    case 23:
    case 13:
    case 38:
    case 29:
      executeParameters = [data.data, data.coercionType];
      break;

    case 37:
    case 28:
      executeParameters = [data.coercionType];
      break;

    case 17:
      executeParameters = [data.subject];
      break;

    case 15:
      executeParameters = [data.recipientField];
      break;

    case 22:
    case 21:
      executeParameters = [data.recipientField, convertComposeEmailDictionaryParameterForSetApi(data.recipientArray)];
      break;

    case 19:
      executeParameters = [data.itemId, data.name];
      break;

    case 16:
      executeParameters = [data.uri, data.name, data.isInline];
      break;

    case 148:
      executeParameters = [data.base64String, data.name, data.isInline];
      break;

    case 20:
      executeParameters = [data.attachmentIndex];
      break;

    case 25:
      executeParameters = [data.TimeProperty, data.time];
      break;

    case 24:
      executeParameters = [data.TimeProperty];
      break;

    case 27:
      executeParameters = [data.location];
      break;

    case 33:
    case 35:
      executeParameters = [data.key, data.type, data.persistent, data.message, data.icon];
      getAdditionalGlobalParametersSingleton().setActionsDefinition(data.actions);
      break;

    case 36:
      executeParameters = [data.key];
      break;

    default:
      optionalParameters = data || {};
      break;
  }

  if (dispid !== 1) {
    executeParameters = getAdditionalGlobalParametersSingleton().updateOutlookExecuteParameters(executeParameters, optionalParameters);
  }

  return executeParameters;
};

var convertRecipientArrayParameterForOutlookForDisplayApi = function convertRecipientArrayParameterForOutlookForDisplayApi(recipients) {
  return recipients != null ? recipients.join(";") : "";
};

var convertComposeEmailDictionaryParameterForSetApi = function convertComposeEmailDictionaryParameterForSetApi(recipients) {
  var results = [];

  if (recipients == null) {
    return results;
  }

  for (var i = 0; i < recipients.length; i++) {
    var newRecipient = [recipients[i].address, recipients[i].name];
    results.push(newRecipient);
  }

  return results;
};

var getItemNumberFromOutlookResponse = function getItemNumberFromOutlookResponse(responseData) {
  var itemNumber = 0;

  if (responseData.length > 2) {
    var extraParameters = JSON.parse(responseData[2]);

    if (!!extraParameters && typeof extraParameters === "object") {
      itemNumber = extraParameters.itemNumber;
    }
  }

  return itemNumber;
};
var createDeserializedData = function createDeserializedData(responseData, itemChanged) {
  var deserializedData = null;
  var returnValues = JSON.parse(responseData[0]);

  if (typeof returnValues === "number") {
    deserializedData = createDeserializedDataWithInt(responseData, itemChanged);
  } else if (!!returnValues && typeof returnValues === "object") {
    deserializedData = createDeserializedDataWithDictionary(responseData, itemChanged);
  } else {
    throw new Error("Return data type from host must be Object or Number");
  }

  return deserializedData;
};

var createDeserializedDataWithDictionary = function createDeserializedDataWithDictionary(responseData, itemChanged) {
  var deserializedData = JSON.parse(responseData[0]);

  if (itemChanged) {
    deserializedData.error = true;
    deserializedData.errorCode = 9030;
  } else if (responseData.length > 1 && responseData[1] !== 0) {
    deserializedData.error = true;
    deserializedData.errorCode = responseData[1];

    if (responseData.length > 2) {
      var diagnosticsData = JSON.parse(responseData[2]);
      deserializedData.diagnostics = diagnosticsData["Diagnostics"];
    }
  } else {
    deserializedData.error = false;
  }

  return deserializedData;
};

var createDeserializedDataWithInt = function createDeserializedDataWithInt(responseData, itemChanged) {
  var deserializedData = {};
  deserializedData.error = true;
  deserializedData.errorCode = responseData[0];
  return deserializedData;
};
// CONCATENATED MODULE: ./src/api/standardInvokeHostMethod.ts







var standardInvokeHostMethod_OSF = __webpack_require__(0);

function standardInvokeHostMethod(dispid, userContext, callback, data, format, customResponse) {
  var executeParameters = callOutlookNativeDispatcher_convertToOutlookNativeParameters(dispid, data);
  CustomFunctionMappings["MailboxHostExecuteApi"](dispid, userContext, callback, executeParameters);
}
function createAsyncResult(value, errorCode, detailedErrorCode, userContext, errorMessage) {
  var initArgs = {};
  var errorArgs;
  initArgs[standardInvokeHostMethod_OSF.DDA.AsyncResultEnum.Properties.Value] = value;
  initArgs[standardInvokeHostMethod_OSF.DDA.AsyncResultEnum.Properties.Context] = userContext;

  if (standardInvokeHostMethod_OSF.DDA.AsyncResultEnum.ErrorCode.Success !== errorCode) {
    errorArgs = {};
    var errorProperties = getErrorArgs(detailedErrorCode);
    errorArgs[standardInvokeHostMethod_OSF.DDA.AsyncResultEnum.ErrorProperties.Name] = errorProperties.name;
    errorArgs[standardInvokeHostMethod_OSF.DDA.AsyncResultEnum.ErrorProperties.Message] = !errorMessage ? errorProperties.message : errorMessage;
    errorArgs[standardInvokeHostMethod_OSF.DDA.AsyncResultEnum.ErrorProperties.Code] = detailedErrorCode;
  }

  return new standardInvokeHostMethod_OSF.DDA.AsyncResult(initArgs, errorArgs);
}
var standardInvokeHostMethod_invokeHostMethod = function invokeHostMethod(dispid, data, responseCallback) {
  var start = performance && performance.now();

  var invokeResponseCallback = function invokeResponseCallback(resultCode, resultData) {
    if (standardInvokeHostMethod_OSF.AppTelemetry) {
      var detailedErroCode = getErrorForTelemetry(resultCode, resultData);
      var end = performance.now();
      standardInvokeHostMethod_OSF.AppTelemetry.onMethodDone(dispid, null, end - start, detailedErroCode);
    }

    if (responseCallback) {
      responseCallback(resultCode, resultData);
    }
  };

  if (standardInvokeHostMethod_OSF.AppName.OutlookWebApp === getAppName()) {
    var args = {
      ApiParams: data,
      MethodData: {
        ControlId: standardInvokeHostMethod_OSF._OfficeAppFactory.getId(),
        DispatchId: dispid
      }
    };

    if (dispid === 1) {
      standardInvokeHostMethod_OSF._OfficeAppFactory.getClientEndPoint().invoke("GetInitialData", invokeResponseCallback, args);
    } else {
      standardInvokeHostMethod_OSF._OfficeAppFactory.getClientEndPoint().invoke("ExecuteMethod", invokeResponseCallback, args);
    }
  } else if (!isOwaOnly(dispid)) {
    callOutlookNativeDispatcher_callOutlookNativeDispatcher(dispid, data, responseCallback);
  } else {
    responseCallback(InvokeResultCode.errorHandlingRequest, null);
  }
};
// CONCATENATED MODULE: ./src/utils/getPermissionLevel.ts

var getPermissionLevel_getPermissionLevel = function getPermissionLevel() {
  return getInitialDataProp("permissionLevel");
};
// CONCATENATED MODULE: ./src/utils/createError.ts
function createError(message, errorInfo) {
  var err = new Error(message);
  err.message = message || "";

  if (errorInfo) {
    for (var v in errorInfo) {
      err[v] = errorInfo[v];
    }
  }

  return err;
}
function createBetaError(featureName) {
  var displayMessage = "The feature {0}, is only enabled on the beta api endpoint".replace("{0}", featureName);
  var err = createError(displayMessage, {
    name: "Sys.FeatureNotEnabled"
  });
  return err;
}
function createParameterCountError(message) {
  var displayMessage = "Sys.ParameterCountException: " + (message ? message : "Parameter count mismatch.");
  var err = createError(displayMessage, {
    name: "Sys.ParameterCountException"
  });
  return err;
}
function createArgumentError(paramName, message) {
  var displayMessage = "Sys.ArgumentException: " + (message ? message : "Value does not fall within the expected range.");

  if (paramName) {
    displayMessage += "\n" + "Parameter name: {0}".replace("{0}", paramName);
  }

  var err = createError(displayMessage, {
    name: "Sys.ArgumentException",
    paramName: paramName
  });
  return err;
}
function createNullArgumentError(paramName, message) {
  var displayMessage = "Sys.ArgumentNullException: " + (message ? message : "Value cannot be null.");

  if (paramName) {
    displayMessage += "\n" + "Parameter name: {0}".replace("{0}", paramName);
  }

  var err = createError(displayMessage, {
    name: "Sys.ArgumentNullException",
    paramName: paramName
  });
  return err;
}
function createArgumentOutOfRange(paramName, actualValue, message) {
  var displayMessage = "Sys.ArgumentOutOfRangeException: " + (message ? message : "Specified argument was out of the range of valid values.");

  if (paramName) {
    displayMessage += "\n" + "Parameter name: {0}".replace("{0}", paramName);
  }

  if (typeof actualValue !== "undefined" && actualValue !== null) {
    displayMessage += "\n" + "Actual value was {0}.".replace("{0}", actualValue);
  }

  var err = createError(displayMessage, {
    name: "Sys.ArgumentOutOfRangeException",
    paramName: paramName,
    actualValue: actualValue
  });
  return err;
}
function createArgumentTypeError(paramName, actualType, expectedType, message) {
  var displayMessage = "Sys.ArgumentTypeException: ";

  if (message) {
    displayMessage += message;
  } else if (actualType && expectedType) {
    displayMessage += "Object of type '{0}' cannot be converted to type '{1}'.".replace("{0}", actualType.getName ? actualType.getName() : actualType).replace("{1}", expectedType.getName ? expectedType.getName() : expectedType);
  } else {
    displayMessage += "Object cannot be converted to the required type.";
  }

  if (paramName) {
    displayMessage += "\n" + "Parameter name: {0}".replace("{0}", paramName);
  }

  var err = createError(displayMessage, {
    name: "Sys.ArgumentTypeException",
    paramName: paramName,
    actualType: actualType,
    expectedType: expectedType
  });
  return err;
}
// CONCATENATED MODULE: ./src/utils/checkPermissionsAndThrow.ts



function checkPermissionsAndThrow(permissions, namespace) {
  if (getPermissionLevel_getPermissionLevel() < permissions) {
    throw createError(getString("l_ElevatedPermissionNeededForMethod_Text").replace("{0}", namespace));
  }
}
// CONCATENATED MODULE: ./src/utils/parseCommonArgs.ts


function parseCommonArgs(args, isCallbackRequired, tryLegacy) {
  var result = {};

  if (tryLegacy) {
    result = tryParseLegacy(args);

    if (result.callback) {
      return result;
    }
  }

  if (args.length === 1) {
    if (typeof args[0] === "function") {
      result.callback = args[0];
    } else if (typeof args[0] === "object") {
      result.options = args[0];
    } else {
      throw createArgumentTypeError();
    }
  } else if (args.length === 2) {
    if (typeof args[0] !== "object") {
      throw createArgumentError("options");
    }

    if (typeof args[1] !== "function") {
      throw createArgumentError("callback");
    }

    result.callback = args[1];
    result.options = args[0];
  } else if (args.length !== 0) {
    throw createParameterCountError(getString("l_ParametersNotAsExpected_Text"));
  }

  if (isCallbackRequired && !result.callback) {
    throw createNullArgumentError("callback");
  }

  if (result.options && result.options.asyncContext) {
    result.asyncContext = result.options.asyncContext;
  }

  return result;
}

function tryParseLegacy(args) {
  var result = {};

  if (args.length === 1 || args.length === 2) {
    if (typeof args[0] !== "function") {
      return result;
    }

    result.callback = args[0];

    if (args.length === 2) {
      result.asyncContext = args[1];
    }

    return result;
  }

  return result;
}
// CONCATENATED MODULE: ./src/validation/recipientConstants.ts
var RecipientFields;

(function (RecipientFields) {
  RecipientFields[RecipientFields["to"] = 0] = "to";
  RecipientFields[RecipientFields["cc"] = 1] = "cc";
  RecipientFields[RecipientFields["bcc"] = 2] = "bcc";
  RecipientFields[RecipientFields["requiredAttendees"] = 0] = "requiredAttendees";
  RecipientFields[RecipientFields["optionalAttendees"] = 1] = "optionalAttendees";
})(RecipientFields || (RecipientFields = {}));

var displayNameLengthLimit = 255;
var recipientsLimit = 100;
var recipientConstants_totalRecipientsLimit = 500;
var maxSmtpLength = 571;
// CONCATENATED MODULE: ./src/validation/displayConstants.ts
var maxLocationLength = 255;
var maxBodyLength = 32 * 1024;
var maxSubjectLength = 255;
var maxRecipients = 100;
var MaxAttachmentNameLength = 255;
var MaxUrlLength = 2048;
var MaxItemIdLength = 200;
var MaxRemoveIdLength = 200;
// CONCATENATED MODULE: ./src/utils/throwOnOutOfRange.ts

function throwOnOutOfRange(value, minValue, maxValue, argumentName) {
  if (value < minValue || value > maxValue) {
    throw createArgumentOutOfRange(String(argumentName));
  }
}
// CONCATENATED MODULE: ./src/utils/OutlookEnums.ts
var MailboxEnums = {};
MailboxEnums.EntityType = {
  MeetingSuggestion: "meetingSuggestion",
  TaskSuggestion: "taskSuggestion",
  Address: "address",
  EmailAddress: "emailAddress",
  Url: "url",
  PhoneNumber: "phoneNumber",
  Contact: "contact",
  FlightReservations: "flightReservations",
  ParcelDeliveries: "parcelDeliveries"
};
MailboxEnums.ItemType = {
  Message: "message",
  Appointment: "appointment"
};
MailboxEnums.ResponseType = {
  None: "none",
  Organizer: "organizer",
  Tentative: "tentative",
  Accepted: "accepted",
  Declined: "declined"
};
MailboxEnums.RecipientType = {
  Other: "other",
  DistributionList: "distributionList",
  User: "user",
  ExternalUser: "externalUser"
};
MailboxEnums.AttachmentType = {
  File: "file",
  Item: "item",
  Cloud: "cloud"
};
MailboxEnums.AttachmentStatus = {
  Added: "added",
  Removed: "removed"
};
MailboxEnums.AttachmentContentFormat = {
  Base64: "base64",
  Url: "url",
  Eml: "eml",
  ICalendar: "iCalendar"
};
MailboxEnums.BodyType = {
  Text: "text",
  Html: "html"
};
MailboxEnums.ItemNotificationMessageType = {
  ProgressIndicator: "progressIndicator",
  InformationalMessage: "informationalMessage",
  ErrorMessage: "errorMessage",
  InsightMessage: "insightMessage"
};
MailboxEnums.Folder = {
  Inbox: "inbox",
  Junk: "junk",
  DeletedItems: "deletedItems"
};
MailboxEnums.ComposeType = {
  Forward: "forward",
  NewMail: "newMail",
  Reply: "reply"
};
var CoercionType = {
  Text: "text",
  Html: "html"
};
MailboxEnums.UserProfileType = {
  Office365: "office365",
  OutlookCom: "outlookCom",
  Enterprise: "enterprise"
};
MailboxEnums.RestVersion = {
  v1_0: "v1.0",
  v2_0: "v2.0",
  Beta: "beta"
};
MailboxEnums.ModuleType = {
  Addins: "addins"
};
MailboxEnums.ActionType = {
  ShowTaskPane: "showTaskPane"
};
MailboxEnums.Days = {
  Mon: "mon",
  Tue: "tue",
  Wed: "wed",
  Thu: "thu",
  Fri: "fri",
  Sat: "sat",
  Sun: "sun",
  Weekday: "weekday",
  WeekendDay: "weekendDay",
  Day: "day"
};
MailboxEnums.WeekNumber = {
  First: "first",
  Second: "second",
  Third: "third",
  Fourth: "fourth",
  Last: "last"
};
MailboxEnums.RecurrenceType = {
  Daily: "daily",
  Weekday: "weekday",
  Weekly: "weekly",
  Monthly: "monthly",
  Yearly: "yearly"
};
MailboxEnums.Month = {
  Jan: "jan",
  Feb: "feb",
  Mar: "mar",
  Apr: "apr",
  May: "may",
  Jun: "jun",
  Jul: "jul",
  Aug: "aug",
  Sep: "sep",
  Oct: "oct",
  Nov: "nov",
  Dec: "dec"
};
MailboxEnums.DelegatePermissions = {
  Read: 0x00000001,
  Write: 0x00000002,
  DeleteOwn: 0x00000004,
  DeleteAll: 0x00000008,
  EditOwn: 0x00000010,
  EditAll: 0x00000020
};
MailboxEnums.TimeZone = {
  AfghanistanStandardTime: "Afghanistan Standard Time",
  AlaskanStandardTime: "Alaskan Standard Time",
  AleutianStandardTime: "Aleutian Standard Time",
  AltaiStandardTime: "Altai Standard Time",
  ArabStandardTime: "Arab Standard Time",
  ArabianStandardTime: "Arabian Standard Time",
  ArabicStandardTime: "Arabic Standard Time",
  ArgentinaStandardTime: "Argentina Standard Time",
  AstrakhanStandardTime: "Astrakhan Standard Time",
  AtlanticStandardTime: "Atlantic Standard Time",
  AUSCentralStandardTime: "AUS Central Standard Time",
  AusCentralWStandardTime: "Aus Central W. Standard Time",
  AUSEasternStandardTime: "AUS Eastern Standard Time",
  AzerbaijanStandardTime: "Azerbaijan Standard Time",
  AzoresStandardTime: "Azores Standard Time",
  BahiaStandardTime: "Bahia Standard Time",
  BangladeshStandardTime: "Bangladesh Standard Time",
  BelarusStandardTime: "Belarus Standard Time",
  BougainvilleStandardTime: "Bougainville Standard Time",
  CanadaCentralStandardTime: "Canada Central Standard Time",
  CapeVerdeStandardTime: "Cape Verde Standard Time",
  CaucasusStandardTime: "Caucasus Standard Time",
  CenAustraliaStandardTime: "Cen. Australia Standard Time",
  CentralAmericaStandardTime: "Central America Standard Time",
  CentralAsiaStandardTime: "Central Asia Standard Time",
  CentralBrazilianStandardTime: "Central Brazilian Standard Time",
  CentralEuropeStandardTime: "Central Europe Standard Time",
  CentralEuropeanStandardTime: "Central European Standard Time",
  CentralPacificStandardTime: "Central Pacific Standard Time",
  CentralStandardTime: "Central Standard Time",
  CentralStandardTime_Mexico: "Central Standard Time (Mexico)",
  ChathamIslandsStandardTime: "Chatham Islands Standard Time",
  ChinaStandardTime: "China Standard Time",
  CubaStandardTime: "Cuba Standard Time",
  DatelineStandardTime: "Dateline Standard Time",
  EAfricaStandardTime: "E. Africa Standard Time",
  EAustraliaStandardTime: "E. Australia Standard Time",
  EEuropeStandardTime: "E. Europe Standard Time",
  ESouthAmericaStandardTime: "E. South America Standard Time",
  EasterIslandStandardTime: "Easter Island Standard Time",
  EasternStandardTime: "Eastern Standard Time",
  EasternStandardTime_Mexico: "Eastern Standard Time (Mexico)",
  EgyptStandardTime: "Egypt Standard Time",
  EkaterinburgStandardTime: "Ekaterinburg Standard Time",
  FijiStandardTime: "Fiji Standard Time",
  FLEStandardTime: "FLE Standard Time",
  GeorgianStandardTime: "Georgian Standard Time",
  GMTStandardTime: "GMT Standard Time",
  GreenlandStandardTime: "Greenland Standard Time",
  GreenwichStandardTime: "Greenwich Standard Time",
  GTBStandardTime: "GTB Standard Time",
  HaitiStandardTime: "Haiti Standard Time",
  HawaiianStandardTime: "Hawaiian Standard Time",
  IndiaStandardTime: "India Standard Time",
  IranStandardTime: "Iran Standard Time",
  IsraelStandardTime: "Israel Standard Time",
  JordanStandardTime: "Jordan Standard Time",
  KaliningradStandardTime: "Kaliningrad Standard Time",
  KamchatkaStandardTime: "Kamchatka Standard Time",
  KoreaStandardTime: "Korea Standard Time",
  LibyaStandardTime: "Libya Standard Time",
  LineIslandsStandardTime: "Line Islands Standard Time",
  LordHoweStandardTime: "Lord Howe Standard Time",
  MagadanStandardTime: "Magadan Standard Time",
  MagallanesStandardTime: "Magallanes Standard Time",
  MarquesasStandardTime: "Marquesas Standard Time",
  MauritiusStandardTime: "Mauritius Standard Time",
  MidAtlanticStandardTime: "Mid-Atlantic Standard Time",
  MiddleEastStandardTime: "Middle East Standard Time",
  MontevideoStandardTime: "Montevideo Standard Time",
  MoroccoStandardTime: "Morocco Standard Time",
  MountainStandardTime: "Mountain Standard Time",
  MountainStandardTime_Mexico: "Mountain Standard Time (Mexico)",
  MyanmarStandardTime: "Myanmar Standard Time",
  NCentralAsiaStandardTime: "N. Central Asia Standard Time",
  NamibiaStandardTime: "Namibia Standard Time",
  NepalStandardTime: "Nepal Standard Time",
  NewZealandStandardTime: "New Zealand Standard Time",
  NewfoundlandStandardTime: "Newfoundland Standard Time",
  NorfolkStandardTime: "Norfolk Standard Time",
  NorthAsiaEastStandardTime: "North Asia East Standard Time",
  NorthAsiaStandardTime: "North Asia Standard Time",
  NorthKoreaStandardTime: "North Korea Standard Time",
  OmskStandardTime: "Omsk Standard Time",
  PacificSAStandardTime: "Pacific SA Standard Time",
  PacificStandardTime: "Pacific Standard Time",
  PacificStandardTime_Mexico: "Pacific Standard Time (Mexico)",
  PakistanStandardTime: "Pakistan Standard Time",
  ParaguayStandardTime: "Paraguay Standard Time",
  RomanceStandardTime: "Romance Standard Time",
  RussiaTimeZone10: "Russia Time Zone 10",
  RussiaTimeZone11: "Russia Time Zone 11",
  RussiaTimeZone3: "Russia Time Zone 3",
  RussianStandardTime: "Russian Standard Time",
  SAEasternStandardTime: "SA Eastern Standard Time",
  SAPacificStandardTime: "SA Pacific Standard Time",
  SAWesternStandardTime: "SA Western Standard Time",
  SaintPierreStandardTime: "Saint Pierre Standard Time",
  SakhalinStandardTime: "Sakhalin Standard Time",
  SamoaStandardTime: "Samoa Standard Time",
  SaratovStandardTime: "Saratov Standard Time",
  SEAsiaStandardTime: "SE Asia Standard Time",
  SingaporeStandardTime: "Singapore Standard Time",
  SouthAfricaStandardTime: "South Africa Standard Time",
  SriLankaStandardTime: "Sri Lanka Standard Time",
  SudanStandardTime: "Sudan Standard Time",
  SyriaStandardTime: "Syria Standard Time",
  TaipeiStandardTime: "Taipei Standard Time",
  TasmaniaStandardTime: "Tasmania Standard Time",
  TocantinsStandardTime: "Tocantins Standard Time",
  TokyoStandardTime: "Tokyo Standard Time",
  TomskStandardTime: "Tomsk Standard Time",
  TongaStandardTime: "Tonga Standard Time",
  TransbaikalStandardTime: "Transbaikal Standard Time",
  TurkeyStandardTime: "Turkey Standard Time",
  TurksAndCaicosStandardTime: "Turks And Caicos Standard Time",
  UlaanbaatarStandardTime: "Ulaanbaatar Standard Time",
  USEasternStandardTime: "US Eastern Standard Time",
  USMountainStandardTime: "US Mountain Standard Time",
  UTC: "UTC",
  UTCPLUS12: "UTC+12",
  UTCPLUS13: "UTC+13",
  UTCMINUS02: "UTC-02",
  UTCMINUS08: "UTC-08",
  UTCMINUS09: "UTC-09",
  UTCMINUS11: "UTC-11",
  VenezuelaStandardTime: "Venezuela Standard Time",
  VladivostokStandardTime: "Vladivostok Standard Time",
  WAustraliaStandardTime: "W. Australia Standard Time",
  WCentralAfricaStandardTime: "W. Central Africa Standard Time",
  WEuropeStandardTime: "W. Europe Standard Time",
  WMongoliaStandardTime: "W. Mongolia Standard Time",
  WestAsiaStandardTime: "West Asia Standard Time",
  WestBankStandardTime: "West Bank Standard Time",
  WestPacificStandardTime: "West Pacific Standard Time",
  YakutskStandardTime: "Yakutsk Standard Time"
};
MailboxEnums.LocationType = {
  Custom: "custom",
  Room: "room"
};
MailboxEnums.AppointmentSensitivityType = {
  Normal: "normal",
  Personal: "personal",
  Private: "private",
  Confidential: "confidential"
};
MailboxEnums.CategoryColor = {
  None: "None",
  Preset0: "Preset0",
  Preset1: "Preset1",
  Preset2: "Preset2",
  Preset3: "Preset3",
  Preset4: "Preset4",
  Preset5: "Preset5",
  Preset6: "Preset6",
  Preset7: "Preset7",
  Preset8: "Preset8",
  Preset9: "Preset9",
  Preset10: "Preset10",
  Preset11: "Preset11",
  Preset12: "Preset12",
  Preset13: "Preset13",
  Preset14: "Preset14",
  Preset15: "Preset15",
  Preset16: "Preset16",
  Preset17: "Preset17",
  Preset18: "Preset18",
  Preset19: "Preset19",
  Preset20: "Preset20",
  Preset21: "Preset21",
  Preset22: "Preset22",
  Preset23: "Preset23",
  Preset24: "Preset24"
};
// CONCATENATED MODULE: ./src/utils/throwOnInvalidRestVersion.ts


function throwOnInvalidRestVersion(restVersion) {
  if (restVersion === null || restVersion === undefined) {
    throw createNullArgumentError(restVersion);
  }

  if (restVersion !== MailboxEnums.RestVersion.v1_0 && restVersion !== MailboxEnums.RestVersion.v2_0 && restVersion !== MailboxEnums.RestVersion.Beta) {
    throw createArgumentError(restVersion);
  }
}
// CONCATENATED MODULE: ./src/utils/convertToRestId.ts


function convertToRestId(itemId, restVersion) {
  if (itemId === null || itemId === undefined) {
    throw createNullArgumentError(itemId);
  }

  throwOnInvalidRestVersion(restVersion);
  return itemId.replace(new RegExp("[/]", "g"), "-").replace(new RegExp("[+]", "g"), "_");
}
// CONCATENATED MODULE: ./src/utils/convertToEwsId.ts


function convertToEwsId(itemId, restVersion) {
  if (itemId === null || itemId === undefined) {
    throw createNullArgumentError(itemId);
  }

  throwOnInvalidRestVersion(restVersion);
  return itemId.replace(new RegExp("[-]", "g"), "/").replace(new RegExp("[_]", "g"), "+");
}
// CONCATENATED MODULE: ./src/validation/validateDisplayForms.ts









function validateRecipientEmails(emailset, name) {
  if (!Array.isArray(emailset)) {
    throw createArgumentTypeError("name");
  }

  throwOnOutOfRange(emailset.length, 0, maxRecipients, "{0}.length".replace("{0}", name));
}
function normalizeRecipientEmails(emailset, name) {
  var originalAttendees = emailset;
  var updatedAttendees = [];

  for (var i = 0; i < originalAttendees.length; i++) {
    if (typeof originalAttendees[i] === "object") {
      throwOnInvalidEmailAddressDetails(originalAttendees[i]);
      updatedAttendees[i] = originalAttendees[i].emailAddress;

      if (typeof updatedAttendees[i] !== "string") {
        throw createArgumentError("{0}[{1}]".replace(name, String(i)));
      }
    } else {
      if (!(typeof originalAttendees[i] === "string")) {
        throw createArgumentError("{0}[{1}]".replace(name, String(i)));
      }

      updatedAttendees[i] = originalAttendees[i];
    }
  }

  return updatedAttendees;
}
function throwOnInvalidEmailAddressDetails(originalAttendee) {
  if (!isNullOrUndefined(originalAttendee.displayName)) {
    if (typeof originalAttendee.displayName === "string" && originalAttendee.displayName.length > displayNameLengthLimit) {
      throw createArgumentOutOfRange("displayName");
    }
  }

  if (!isNullOrUndefined(originalAttendee.emailAddress)) {
    if (typeof originalAttendee.emailAddress === "string" && originalAttendee.emailAddress.length > maxSmtpLength) {
      throw createArgumentOutOfRange("emailAddress");
    }
  }

  if (!isNullOrUndefined(originalAttendee.appointmentResponse)) {
    if (typeof originalAttendee.appointmentResponse !== "string") {
      throw createArgumentOutOfRange("appointmentResponse");
    }
  }

  if (!isNullOrUndefined(originalAttendee.recipientType)) {
    if (typeof originalAttendee.recipientType !== "string") {
      throw createArgumentOutOfRange("recipientType");
    }
  }
}
function validateDisplayFormParameters(itemId) {
  if (typeof itemId === "string") {
    throwOnInvalidItemId(itemId);
  } else {
    throw createArgumentTypeError("itemId");
  }
}

function throwOnInvalidItemId(itemId) {
  if (isNullOrUndefined(itemId) || itemId === "") {
    throw createNullArgumentError("itemId");
  }
}

function getItemIdBasedOnHost(itemId) {
  if (getInitialDataProp("isRestIdSupported")) {
    return convertToRestId(itemId, MailboxEnums.RestVersion.v1_0);
  }

  return convertToEwsId(itemId, MailboxEnums.RestVersion.v1_0);
}
// CONCATENATED MODULE: ./src/utils/RuntimeFlighting.ts

var beta = 2;
var production = 1;
var currentLevel;
currentLevel = production;
var getCurrentLevel = function getCurrentLevel() {
  return currentLevel;
};
var Features = {
  featureSampleProduction: production,
  featureSampleBeta: beta,
  calendarItems: beta,
  signature: beta,
  replyCallback: beta,
  displayXAsync: beta,
  propertyGetAll: beta,
  state: beta
};
function isFeatureEnabled(feature) {
  return feature <= getCurrentLevel();
}
function checkFeatureEnabledAndThrow(feature, featureName) {
  if (!isFeatureEnabled(feature)) {
    throw createBetaError(featureName);
  }
}
// CONCATENATED MODULE: ./src/methods/displayAppointmentForm.ts
var __spreadArrays = undefined && undefined.__spreadArrays || function () {
  for (var s = 0, i = 0, il = arguments.length; i < il; i++) {
    s += arguments[i].length;
  }

  for (var r = Array(s), k = 0, i = 0; i < il; i++) {
    for (var a = arguments[i], j = 0, jl = a.length; j < jl; j++, k++) {
      r[k] = a[j];
    }
  }

  return r;
};






function displayAppointmentForm(itemId) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  displayAppointmentFormHelper.apply(void 0, __spreadArrays([9, itemId], args));
}
function displayAppointmentFormAsync(itemId) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkFeatureEnabledAndThrow(Features.displayXAsync, "displayAppointmentFormAsync");
  displayAppointmentFormHelper.apply(void 0, __spreadArrays([180, itemId], args));
}

function displayAppointmentFormHelper(dispidToInvoke, itemId) {
  var args = [];

  for (var _i = 2; _i < arguments.length; _i++) {
    args[_i - 2] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "mailbox.displayAppointmentForm");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    itemId: itemId
  };
  validateParameters(parameters);
  standardInvokeHostMethod(dispidToInvoke, commonParameters.asyncContext, commonParameters.callback, {
    itemId: getItemIdBasedOnHost(parameters.itemId)
  }, undefined);
}

function validateParameters(parameters) {
  validateDisplayFormParameters(parameters.itemId);
}
// CONCATENATED MODULE: ./src/methods/displayMessageForm.ts
var displayMessageForm_spreadArrays = undefined && undefined.__spreadArrays || function () {
  for (var s = 0, i = 0, il = arguments.length; i < il; i++) {
    s += arguments[i].length;
  }

  for (var r = Array(s), k = 0, i = 0; i < il; i++) {
    for (var a = arguments[i], j = 0, jl = a.length; j < jl; j++, k++) {
      r[k] = a[j];
    }
  }

  return r;
};






function displayMessageForm(itemId) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  displayMessageFormHelper.apply(void 0, displayMessageForm_spreadArrays([8, itemId], args));
}
function displayMessageFormAsync(itemId) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkFeatureEnabledAndThrow(Features.displayXAsync, "displayMessageFormAsync");
  displayMessageFormHelper.apply(void 0, displayMessageForm_spreadArrays([179, itemId], args));
}

function displayMessageFormHelper(dispidToInvoke, itemId) {
  var args = [];

  for (var _i = 2; _i < arguments.length; _i++) {
    args[_i - 2] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "mailbox.displayMessageForm");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    itemId: itemId
  };
  displayMessageForm_validateParameters(parameters);
  standardInvokeHostMethod(dispidToInvoke, commonParameters.asyncContext, commonParameters.callback, {
    itemId: getItemIdBasedOnHost(parameters.itemId)
  }, undefined);
}

function displayMessageForm_validateParameters(parameters) {
  validateDisplayFormParameters(parameters.itemId);
}
// CONCATENATED MODULE: ./src/utils/validateOptionalStringParameter.ts


function validateOptionalStringParameter(value, minLength, maxlength, name) {
  if (typeof value === "string") {
    throwOnOutOfRange(value.length, minLength, maxlength, name);
  } else {
    throw createArgumentError(String(name));
  }
}
// CONCATENATED MODULE: ./src/methods/displayNewAppointmentForm.ts
var displayNewAppointmentForm_spreadArrays = undefined && undefined.__spreadArrays || function () {
  for (var s = 0, i = 0, il = arguments.length; i < il; i++) {
    s += arguments[i].length;
  }

  for (var r = Array(s), k = 0, i = 0; i < il; i++) {
    for (var a = arguments[i], j = 0, jl = a.length; j < jl; j++, k++) {
      r[k] = a[j];
    }
  }

  return r;
};











function displayNewAppointmentForm(parameters) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  return displayNewAppointmentFormHelper.apply(void 0, displayNewAppointmentForm_spreadArrays([7, parameters], args));
}
function displayNewAppointmentFormAsync(parameters) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkFeatureEnabledAndThrow(Features.displayXAsync, "displayNewAppointmentFormAsync");
  return displayNewAppointmentFormHelper.apply(void 0, displayNewAppointmentForm_spreadArrays([177, parameters], args));
}

function displayNewAppointmentFormHelper(dispidToInvoke, parameters) {
  var args = [];

  for (var _i = 2; _i < arguments.length; _i++) {
    args[_i - 2] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "mailbox.displayNewAppointmentForm");
  var commonParameters = parseCommonArgs(args, false, false);
  displayNewAppointmentForm_validateParameters(parameters);
  var updatedParameters = normalizeParameters(parameters);
  standardInvokeHostMethod(dispidToInvoke, commonParameters.asyncContext, commonParameters.callback, updatedParameters, undefined);
}

function displayNewAppointmentForm_validateParameters(parameters) {
  if (!isNullOrUndefined(parameters.requiredAttendees)) {
    validateRecipientEmails(parameters.requiredAttendees, "requiredAttendees");
  }

  if (!isNullOrUndefined(parameters.optionalAttendees)) {
    validateRecipientEmails(parameters.optionalAttendees, "optionalAttendees");
  }

  if (!isNullOrUndefined(parameters.location)) {
    validateOptionalStringParameter(parameters.location, 0, maxLocationLength, "location");
  }

  if (!isNullOrUndefined(parameters.body)) {
    validateOptionalStringParameter(parameters.body, 0, maxBodyLength, "body");
  }

  if (!isNullOrUndefined(parameters.subject)) {
    validateOptionalStringParameter(parameters.subject, 0, maxSubjectLength, "subject");
  }

  if (!isNullOrUndefined(parameters.start)) {
    if (!(parameters.start instanceof Date)) {
      throw createArgumentError("start");
    }

    if (!isNullOrUndefined(parameters.end)) {
      if (!(parameters.end instanceof Date)) {
        throw createArgumentError("end");
      }

      if (parameters.end < parameters.start) {
        throw createArgumentError("end", getString("l_InvalidEventDates_Text"));
      }
    }
  }
}

function normalizeParameters(parameters) {
  var normalizedRequiredAttendees = null;
  var normalizedOptionalAttendees = null;

  if (!isNullOrUndefined(parameters.requiredAttendees)) {
    normalizedRequiredAttendees = normalizeRecipientEmails(parameters.requiredAttendees, "requiredAttendees");
  }

  if (!isNullOrUndefined(parameters.optionalAttendees)) {
    normalizedOptionalAttendees = normalizeRecipientEmails(parameters.optionalAttendees, "optionalAttendees");
  }

  if (!isNullOrUndefined(parameters.start)) {
    var startDate = parameters.start;
    parameters.start = startDate.getTime();
  }

  if (!isNullOrUndefined(parameters.end)) {
    var endDate = parameters.end;
    parameters.end = endDate.getTime();
  }

  var updatedParameters = JSON.parse(JSON.stringify(parameters));

  if (normalizedRequiredAttendees || normalizedOptionalAttendees) {
    if (!isNullOrUndefined(parameters.requiredAttendees)) {
      updatedParameters.requiredAttendees = normalizedRequiredAttendees;
    }

    if (!isNullOrUndefined(parameters.optionalAttendees)) {
      updatedParameters.optionalAttendees = normalizedOptionalAttendees;
    }
  }

  return updatedParameters;
}
// CONCATENATED MODULE: ./src/methods/displayNewMessageForm.ts
var displayNewMessageForm_spreadArrays = undefined && undefined.__spreadArrays || function () {
  for (var s = 0, i = 0, il = arguments.length; i < il; i++) {
    s += arguments[i].length;
  }

  for (var r = Array(s), k = 0, i = 0; i < il; i++) {
    for (var a = arguments[i], j = 0, jl = a.length; j < jl; j++, k++) {
      r[k] = a[j];
    }
  }

  return r;
};












function displayNewMessageForm(parameters) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  return displayNewMessageFormHelper.apply(void 0, displayNewMessageForm_spreadArrays([44, parameters], args));
}
function displayNewMessageFormAsync(parameters) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkFeatureEnabledAndThrow(Features.displayXAsync, "displayNewMessageFormAsync");
  return displayNewMessageFormHelper.apply(void 0, displayNewMessageForm_spreadArrays([178, parameters], args));
}

function displayNewMessageFormHelper(dispidToInvoke, parameters) {
  var args = [];

  for (var _i = 2; _i < arguments.length; _i++) {
    args[_i - 2] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "mailbox.displayNewMessageForm");
  var commonParameters = parseCommonArgs(args, false, false);
  displayNewMessageForm_validateParameters(parameters);
  var updatedParameters = normailzeParameters(parameters);
  standardInvokeHostMethod(dispidToInvoke, commonParameters.asyncContext, commonParameters.callback, updatedParameters === null || updatedParameters === undefined ? parameters : updatedParameters, undefined);
}

function displayNewMessageForm_validateParameters(parameters) {
  if (parameters !== null && parameters !== null) {
    if (!isNullOrUndefined(parameters.toRecipients)) {
      validateRecipientEmails(parameters.toRecipients, "toRecipients");
    }

    if (!isNullOrUndefined(parameters.ccRecipients)) {
      validateRecipientEmails(parameters.ccRecipients, "ccRecipients");
    }

    if (!isNullOrUndefined(parameters.bccRecipients)) {
      validateRecipientEmails(parameters.bccRecipients, "bccRecipients");
    }

    if (!isNullOrUndefined(parameters.htmlBody)) {
      validateOptionalStringParameter(parameters.htmlBody, 0, maxBodyLength, "htmlBody");
    }

    if (!isNullOrUndefined(parameters.subject)) {
      validateOptionalStringParameter(parameters.subject, 0, maxSubjectLength, "subject");
    }
  }
}

function normailzeParameters(parameters) {
  var updatedParameters = JSON.parse(JSON.stringify(parameters));

  if (!isNullOrUndefined(parameters)) {
    if (parameters.toRecipients) {
      updatedParameters.toRecipients = normalizeRecipientEmails(parameters.toRecipients, "toRecipients");
    }

    if (parameters.ccRecipients) {
      updatedParameters.ccRecipients = normalizeRecipientEmails(parameters.ccRecipients, "ccRecipients");
    }

    if (parameters.bccRecipients) {
      updatedParameters.bccRecipients = normalizeRecipientEmails(parameters.bccRecipients, "bccRecipients");
    }

    var attachments = getAttachments(parameters);

    if (parameters.attachments) {
      updatedParameters.attachments = createAttachmentsDataForHost(attachments);
    }
  }

  return updatedParameters;
}
function getAttachments(data) {
  var attachments = [];

  if (data.attachments) {
    attachments = data.attachments;
    throwOnInvalidAttachmentsArray(attachments);
  }

  return attachments;
}
function throwOnInvalidAttachmentsArray(attachments) {
  if (!isNullOrUndefined(attachments) && !Array.isArray(attachments)) {
    throw createArgumentError("attachments");
  }
}
function createAttachmentsDataForHost(attachments) {
  var attachmentsData = [];

  for (var i = 0; i < attachments.length; i++) {
    if (typeof attachments[i] === "object") {
      var attachment = attachments[i];
      throwOnInvalidAttachment(attachment);
      attachmentsData.push(createAttachmentData(attachment));
    } else {
      throw createArgumentError("attachments");
    }
  }

  return attachmentsData;
}
function throwOnInvalidAttachment(attachment) {
  if (typeof attachment !== "object") {
    throw createArgumentError("attachments");
  }

  if (!attachment.type || !attachment.name) {
    throw createArgumentError("attachments");
  }

  if (!attachment.url && !attachment.itemId) {
    throw createArgumentError("attachments");
  }
}
function createAttachmentData(attachment) {
  var attachmentData = null;

  if (attachment.type === MailboxEnums.AttachmentType.File) {
    var url = attachment.url;
    var name_1 = attachment.name;
    var isInline = !!attachment.isInline;
    throwOnInvalidAttachmentUrlOrName(url, name_1);
    attachmentData = [MailboxEnums.AttachmentType.File, name_1, url, isInline];
  } else if (attachment.type === MailboxEnums.AttachmentType.Item) {
    var itemId = getItemIdBasedOnHost(attachment.itemId);
    var name_2 = attachment.name;
    throwOnInvalidAttachmentItemIdOrName(itemId, name_2);
    attachmentData = [MailboxEnums.AttachmentType.Item, name_2, itemId];
  } else {
    throw createArgumentError("attachments");
  }

  return attachmentData;
}
function throwOnInvalidAttachmentUrlOrName(url, name) {
  if (!(typeof url === "string") && !(typeof name === "string")) {
    throw createArgumentError("attachments");
  }

  if (url.length > MaxUrlLength) {
    throw createArgumentOutOfRange("attachments", url.length, getString("l_AttachmentUrlTooLong_Text"));
  }

  throwOnInvalidAttachmentName(name);
}
function throwOnInvalidAttachmentName(name) {
  if (name.length > MaxAttachmentNameLength) {
    throw createArgumentOutOfRange("attachments", name.length, getString("l_AttachmentNameTooLong_Text"));
  }
}
function throwOnInvalidAttachmentItemIdOrName(itemId, name) {
  if (!(typeof itemId === "string") || !(typeof name === "string")) {
    throw createArgumentError("attachments");
  }

  if (itemId.length > MaxItemIdLength) {
    throw createArgumentOutOfRange("attachments", itemId.length, getString("l_AttachmentItemIdTooLong_Text"));
  }

  throwOnInvalidAttachmentName(name);
}
// CONCATENATED MODULE: ./src/utils/handleTokenResponse.ts





var handleTokenResponse_OSF = __webpack_require__(0);

function handleTokenResponse(response, context, resultCode) {
  var asyncResult = undefined;

  if (!!resultCode && resultCode !== InvokeResultCode.noError) {
    asyncResult = createAsyncResult(undefined, handleTokenResponse_OSF.DDA.AsyncResultEnum.ErrorCode.Failed, 9017, context, getString("l_InternalProtocolError_Text").replace("{0}", resultCode));

    if (!!asyncResult) {
      asyncResult.diagnostics = {
        InvokeCodeResult: resultCode
      };
    }
  } else {
    if (getAppName() === handleTokenResponse_OSF.AppName.Outlook && response.error !== undefined && response.errorCode !== undefined && !!response.error && response.errorCode === 9030) {
      asyncResult = createAsyncResult(undefined, handleTokenResponse_OSF.DDA.AsyncResultEnum.ErrorCode.Failed, response.errorCode, context, response.errorMessage);
    } else if (!!response.wasSuccessful) {
      asyncResult = createAsyncResult(response.token, handleTokenResponse_OSF.DDA.AsyncResultEnum.ErrorCode.Success, 0, context);
    } else {
      asyncResult = createAsyncResult(undefined, handleTokenResponse_OSF.DDA.AsyncResultEnum.ErrorCode.Failed, response.errorCode, context, response.errorMessage);
    }

    if (response.diagnostics) {
      asyncResult.diagnostics = response.diagnostics;
    }
  }

  return asyncResult;
}
// CONCATENATED MODULE: ./src/methods/getCallbackToken.ts








function getCallbackToken() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "mailbox.getCallbackTokenAsync");
  var commonParameters = parseCommonArgs(args, true, true);
  var isRest = false;

  if (commonParameters.options && !!commonParameters.options.isRest) {
    isRest = true;
  }

  if (getIsNoItemContextWebExt()) {
    if (!isRest || getPermissionLevel_getPermissionLevel() < 3) {
      throw createError(getString("l_TokenAccessDeniedWithoutItemContext_Text"));
    }
  }

  standardInvokeHostMethod(12, commonParameters.asyncContext, commonParameters.callback, {
    isRest: isRest
  }, undefined, handleTokenResponse);
}
// CONCATENATED MODULE: ./src/methods/getUserIdentityToken.ts




function getUserIdentityToken() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "mailbox.getUserIdentityToken");
  var commonParameters = parseCommonArgs(args, true, true);
  standardInvokeHostMethod(2, commonParameters.asyncContext, commonParameters.callback, undefined, undefined, handleTokenResponse);
}
// CONCATENATED MODULE: ./src/methods/makeEwsRequest.ts







var makeEwsRequest_OSF = __webpack_require__(0);

var maxEwsRequestSize = 1000000;
function makeEwsRequest(body) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(3, "mailbox.makeEwsRequest");
  var commonParameters = parseCommonArgs(args, true, true);

  if (body === null || body === undefined) {
    throw createNullArgumentError("data");
  }

  if (typeof body !== "string") {
    throw createArgumentTypeError("data", typeof body, "string");
  }

  if (body.length > maxEwsRequestSize) {
    throw createArgumentError("data", getString("l_EwsRequestOversized_Text"));
  }

  standardInvokeHostMethod(5, commonParameters.asyncContext, commonParameters.callback, {
    body: body
  }, undefined, handleCustomResponse);
}

function handleCustomResponse(data, context, responseCode) {
  if (!!responseCode && responseCode !== InvokeResultCode.noError) {
    return createAsyncResult(undefined, makeEwsRequest_OSF.DDA.AsyncResultEnum.ErrorCode.Failed, 9017, context, getString("l_InternalProtocolError_Text").replace("{0}", responseCode));
  } else if (data.wasProxySuccessful === false) {
    return createAsyncResult(undefined, makeEwsRequest_OSF.DDA.AsyncResultEnum.ErrorCode.Failed, 9020, context, data.errorMessage);
  } else {
    return createAsyncResult(data.body, makeEwsRequest_OSF.DDA.AsyncResultEnum.ErrorCode.Success, 0, context);
  }
}
// CONCATENATED MODULE: ./src/utils/objectDefine.ts
var objectDefine = function objectDefine(o, props) {
  var keys = Object.keys(props);
  var values = keys.map(function (prop) {
    return {
      value: props[prop],
      writable: false
    };
  });
  var properties = {};
  keys.forEach(function (key, index) {
    properties[key] = values[index];
  });
  return Object.defineProperties(o, properties);
};
// CONCATENATED MODULE: ./src/api/getDiagnostics.ts



var getDiagnostics_OSF = __webpack_require__(0);

var getHostName = function getHostName() {
  return "Outlook";
};
function getDiagnosticsSurface() {
  return objectDefine({}, {
    hostName: getHostName(),
    hostVersion: getInitialDataProp("hostVersion"),
    OWAView: getInitialDataProp("owaView")
  });
}
// CONCATENATED MODULE: ./src/api/getUserProfile.ts


function getUserProfileSurface() {
  return objectDefine({}, {
    accountType: getInitialDataProp("userProfileType"),
    displayName: getInitialDataProp("userDisplayName"),
    emailAddress: getInitialDataProp("userEmailAddress"),
    timeZone: getInitialDataProp("userTimeZone")
  });
}
// CONCATENATED MODULE: ./src/validation/categoryConstants.ts

var CategoryColor = MailboxEnums.CategoryColor;
var categoriesCharacterLimit = 255;
var colorPresets = [CategoryColor.None, CategoryColor.Preset0, CategoryColor.Preset1, CategoryColor.Preset2, CategoryColor.Preset3, CategoryColor.Preset4, CategoryColor.Preset5, CategoryColor.Preset6, CategoryColor.Preset7, CategoryColor.Preset8, CategoryColor.Preset9, CategoryColor.Preset10, CategoryColor.Preset11, CategoryColor.Preset12, CategoryColor.Preset13, CategoryColor.Preset14, CategoryColor.Preset15, CategoryColor.Preset16, CategoryColor.Preset17, CategoryColor.Preset18, CategoryColor.Preset19, CategoryColor.Preset20, CategoryColor.Preset21, CategoryColor.Preset22, CategoryColor.Preset23, CategoryColor.Preset24];
// CONCATENATED MODULE: ./src/validation/validateCategoryDetailsArray.ts


function validateCategoryDetailsArray(categoryDetails) {
  if (!categoryDetails) {
    throw createArgumentError("categoryDetails");
  }

  if (!Array.isArray(categoryDetails)) {
    throw createArgumentTypeError("categoryDetails", typeof categoryDetails, typeof []);
  }

  if (categoryDetails.length === 0) {
    throw createArgumentError("categoryDetails");
  }

  categoryDetails.forEach(validateCategoryDetails);
}

function validateCategoryDetails(categoryDetails) {
  if (!categoryDetails) {
    throw createArgumentError("categoryDetails");
  }

  if (!categoryDetails.color || !categoryDetails.displayName) {
    throw createArgumentError("categoryDetails");
  }

  if (typeof categoryDetails.color !== "string") {
    throw createArgumentTypeError("categoryDetails.color", typeof categoryDetails.color, "string");
  }

  if (typeof categoryDetails.displayName !== "string") {
    throw createArgumentTypeError("categoryDetails.displayName", typeof categoryDetails.displayName, "string");
  }

  if (categoryDetails.displayName.length > categoriesCharacterLimit) {
    throw createArgumentOutOfRange("categoryDetails.displayName", categoryDetails.displayName.length);
  }

  if (colorPresets.indexOf(categoryDetails.color) === -1) {
    throw createArgumentError("categoryDetails.color");
  }
}
// CONCATENATED MODULE: ./src/methods/addMasterCategories.ts




function addMasterCategories(categoryDetails) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(3, "masterCategories.addAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    categoryDetails: categoryDetails
  };
  validateCategoryDetailsArray(categoryDetails);
  standardInvokeHostMethod(161, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}
// CONCATENATED MODULE: ./src/methods/getMasterCategories.ts



function getMasterCategories() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(3, "masterCategories.getAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  standardInvokeHostMethod(160, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/validation/validateCategoryArray.ts


function validateCategoryArray(categories) {
  if (!categories) {
    throw createArgumentError("categories");
  }

  if (!Array.isArray(categories)) {
    throw createArgumentTypeError("categories", typeof categories, typeof Array);
  }

  if (categories.length === 0) {
    throw createArgumentError("categories");
  }

  categories.forEach(validateCategory);
}

function validateCategory(category) {
  if (!category) {
    throw createArgumentError("categories");
  }

  if (typeof category !== "string") {
    throw createArgumentTypeError("categories", typeof category, "string");
  }

  if (category.length > categoriesCharacterLimit) {
    throw createArgumentOutOfRange("categories", category.length);
  }
}
// CONCATENATED MODULE: ./src/methods/removeMasterCategories.ts




function removeMasterCategories(categories) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(3, "masterCategories.removeAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    categories: categories
  };
  validateCategoryArray(categories);
  standardInvokeHostMethod(162, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}
// CONCATENATED MODULE: ./src/api/getMasterCategoriesSurface.ts




function getMasterCategoriesSurface() {
  return objectDefine({}, {
    addAsync: addMasterCategories,
    getAsync: getMasterCategories,
    removeAsync: removeMasterCategories
  });
}
// CONCATENATED MODULE: ./src/methods/closeApp.ts

function closeApp() {
  standardInvokeHostMethod(42, undefined, undefined, undefined, undefined);
}
// CONCATENATED MODULE: ./src/utils/getHostItemType.ts

var getHostItemType_getHostItemType = function getHostItemType() {
  return getInitialDataProp("itemType");
};
// CONCATENATED MODULE: ./src/utils/HostItemType.ts
var HostItemType;

(function (HostItemType) {
  HostItemType[HostItemType["Message"] = 1] = "Message";
  HostItemType[HostItemType["Appointment"] = 2] = "Appointment";
  HostItemType[HostItemType["MeetingRequest"] = 3] = "MeetingRequest";
  HostItemType[HostItemType["MessageCompose"] = 4] = "MessageCompose";
  HostItemType[HostItemType["AppointmentCompose"] = 5] = "AppointmentCompose";
  HostItemType[HostItemType["ItemLess"] = 6] = "ItemLess";
})(HostItemType || (HostItemType = {}));
// CONCATENATED MODULE: ./src/methods/getInitializationContext.ts



function getInitializationContext() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "item.getInitializationContext");
  var commonParameters = parseCommonArgs(args, true, false);
  standardInvokeHostMethod(99, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/validation/customPropertiesConstants.ts
var DatePrefix = "Date(";
var DatePostfix = ")";
var MaxCustomPropertiesLength = 2500;
var CustomPropertyType;

(function (CustomPropertyType) {
  CustomPropertyType[CustomPropertyType["NonTransmittable"] = 0] = "NonTransmittable";
})(CustomPropertyType || (CustomPropertyType = {}));
// CONCATENATED MODULE: ./src/methods/saveCustomProperties.ts





function saveCustomProperties(customProperties) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "item.saveCustomProperties");
  var commonParameters = parseCommonArgs(args, false, true);
  saveCustomProperties_validateParameters(customProperties);
  standardInvokeHostMethod(4, commonParameters.asyncContext, commonParameters.callback, {
    customProperties: customProperties
  }, undefined);
}

function saveCustomProperties_validateParameters(customProperties) {
  if (JSON.stringify(customProperties).length > MaxCustomPropertiesLength) {
    throw createArgumentOutOfRange("customProperties");
  }
}
// CONCATENATED MODULE: ./src/api/CustomProperties.ts
var CustomProperties_spreadArrays = undefined && undefined.__spreadArrays || function () {
  for (var s = 0, i = 0, il = arguments.length; i < il; i++) {
    s += arguments[i].length;
  }

  for (var r = Array(s), k = 0, i = 0; i < il; i++) {
    for (var a = arguments[i], j = 0, jl = a.length; j < jl; j++, k++) {
      r[k] = a[j];
    }
  }

  return r;
};







var CustomProperties_CustomProperties = function () {
  function CustomProperties(deserializedData) {
    if (isNullOrUndefined(deserializedData)) {
      createNullArgumentError("data");
    }

    if (Array.isArray(deserializedData)) {
      var customropertiesArray = deserializedData;

      if (customropertiesArray.length > CustomPropertyType.NonTransmittable) {
        deserializedData = customropertiesArray[CustomPropertyType.NonTransmittable];
      } else {
        throw createArgumentError("data");
      }
    } else {
      this.rawData = deserializedData;
    }
  }

  CustomProperties.prototype.get = function (key) {
    var value = this.rawData[key];

    if (typeof value === "string") {
      var valueString = value;

      if (valueString.length > DatePrefix.length + DatePostfix.length && valueString.startsWith(DatePrefix) && valueString.endsWith(DatePostfix)) {
        var ticksString = valueString.substring(DatePrefix.length, valueString.length - 1);
        var ticks = parseInt(ticksString);

        if (!isNaN(ticks)) {
          var dateTimeValue = new Date(ticks);

          if (!isNullOrUndefined(dateTimeValue)) {
            value = dateTimeValue;
          }
        }
      }
    }

    return value;
  };

  CustomProperties.prototype.set = function (key, value) {
    if (value instanceof Date) {
      value = DatePrefix + value.getTime() + DatePostfix;
    }

    this.rawData[key] = value;
  };

  CustomProperties.prototype.remove = function (key) {
    delete this.rawData[key];
  };

  CustomProperties.prototype.saveAsync = function () {
    var args = [];

    for (var _i = 0; _i < arguments.length; _i++) {
      args[_i] = arguments[_i];
    }

    saveCustomProperties.apply(void 0, CustomProperties_spreadArrays([this.rawData], args));
  };

  CustomProperties.prototype.getAll = function () {
    var _this = this;

    checkFeatureEnabledAndThrow(Features.propertyGetAll, "getAll");
    var dictionary = {};
    var keys = Object.keys(this.rawData);
    keys.forEach(function (key) {
      dictionary[key] = _this.get(key);
    });
    return dictionary;
  };

  return CustomProperties;
}();


// CONCATENATED MODULE: ./src/methods/loadCustomProperties.ts






var loadCustomProperties_OSF = __webpack_require__(0);

function loadCustomProperties() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  var commonParameters = parseCommonArgs(args, true, true);
  standardInvokeHostMethod(3, commonParameters.asyncContext, commonParameters.callback, undefined, undefined, loadCustomProperties_handleCustomResponse);
}

function loadCustomProperties_handleCustomResponse(data, context, responseCode) {
  if (typeof responseCode !== "undefined" && responseCode !== InvokeResultCode.noError) {
    return createAsyncResult(undefined, loadCustomProperties_OSF.DDA.AsyncResultEnum.ErrorCode.Failed, 9017, context, getString("l_InternalProtocolError_Text").replace("{0}", responseCode));
  } else if (data.wasSuccessful) {
    var props = JSON.parse(data.customProperties);
    var value = new CustomProperties_CustomProperties(props);
    return createAsyncResult(value, loadCustomProperties_OSF.DDA.AsyncResultEnum.ErrorCode.Success, 0, context);
  } else {
    return createAsyncResult(undefined, loadCustomProperties_OSF.DDA.AsyncResultEnum.ErrorCode.Failed, 9020, context, data.errorMessage);
  }
}
// CONCATENATED MODULE: ./src/utils/bodyUtils.ts



var bodyUtils_OSF = __webpack_require__(0);

var HostCoercionType;

(function (HostCoercionType) {
  HostCoercionType[HostCoercionType["Text"] = 0] = "Text";
  HostCoercionType[HostCoercionType["Html"] = 3] = "Html";
})(HostCoercionType || (HostCoercionType = {}));

function addCoercionTypeParameter(parameters, args) {
  if (!!args.options && typeof args.options.coercionType === "string") {
    parameters.coercionType = getCoercionTypeFromString(args.options.coercionType);
  } else {
    parameters.coercionType = HostCoercionType.Text;
  }
}
function getCoercionTypeFromString(coercionType) {
  if (coercionType === CoercionType.Html) {
    return HostCoercionType.Html;
  } else if (coercionType === CoercionType.Text) {
    return HostCoercionType.Text;
  } else {
    return undefined;
  }
}
function invokeCallbackWithCoercionTypeError(args) {
  args.callback && args.callback(createAsyncResult(undefined, bodyUtils_OSF.DDA.AsyncResultEnum.ErrorCode.Failed, 1000, args.asyncContext));
}
// CONCATENATED MODULE: ./src/methods/getBody.ts





function getBody(coercionType) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "body.getAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  var parameters = {
    coercionType: getCoercionTypeFromString(coercionType)
  };

  if (parameters.coercionType === undefined) {
    throw createArgumentError("coercionType");
  }

  standardInvokeHostMethod(37, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}
// CONCATENATED MODULE: ./src/methods/getBodyType.ts



function getBodyType() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "body.getTypeAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  standardInvokeHostMethod(14, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/validation/validateBodyParameters.ts

var maxDataLengthForBodyApi = 1000000;
var maxAppendOnSendLength = 5000;
var maxDataLengthForSignatureBodyApi = 30000;
function validateAppendOnSendBodyParamters(parameters) {
  if (typeof parameters.appendTxt !== "string") {
    throw createArgumentTypeError("data", typeof parameters.appendTxt, "string");
  }

  if (parameters.appendTxt.length > maxAppendOnSendLength) {
    throw createArgumentOutOfRange("data", parameters.appendTxt.length);
  }
}
function validateBodyParameters(parameters) {
  if (typeof parameters.data !== "string") {
    throw createArgumentTypeError("data", typeof parameters.data, "string");
  }

  if (parameters.data.length > maxDataLengthForBodyApi) {
    throw createArgumentOutOfRange("data", parameters.data.length);
  }
}
function validateSignatureBodyParameters(parameters) {
  if (typeof parameters.data !== "string") {
    throw createArgumentTypeError("data", typeof parameters.data, "string");
  }

  if (parameters.data.length > maxDataLengthForSignatureBodyApi) {
    throw createArgumentOutOfRange("data", parameters.data.length);
  }
}
// CONCATENATED MODULE: ./src/methods/setBody.ts





function setBody(data) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "body.setAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    data: data
  };
  validateBodyParameters(parameters);
  addCoercionTypeParameter(parameters, commonParameters);

  if (parameters.coercionType === undefined) {
    invokeCallbackWithCoercionTypeError(commonParameters);
    return;
  }

  standardInvokeHostMethod(38, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}
// CONCATENATED MODULE: ./src/methods/bodyPrepend.ts





function bodyPrepend(data) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "body.prependAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    data: data
  };
  validateBodyParameters(parameters);
  addCoercionTypeParameter(parameters, commonParameters);

  if (parameters.coercionType === undefined) {
    invokeCallbackWithCoercionTypeError(commonParameters);
    return;
  }

  standardInvokeHostMethod(23, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}
// CONCATENATED MODULE: ./src/methods/appendOnSend.ts






function appendOnSend(data) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "body.appendOnSendAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    appendTxt: data
  };

  if (isNullOrUndefined(parameters.appendTxt)) {
    parameters.appendTxt = "";
  } else {
    validateAppendOnSendBodyParamters(parameters);
  }

  addCoercionTypeParameter(parameters, commonParameters);

  if (parameters.coercionType === undefined) {
    invokeCallbackWithCoercionTypeError(commonParameters);
    return;
  }

  standardInvokeHostMethod(100, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}
// CONCATENATED MODULE: ./src/methods/setSelectedData.ts





function setSelectedData(dispid) {
  return function (data) {
    var args = [];

    for (var _i = 1; _i < arguments.length; _i++) {
      args[_i - 1] = arguments[_i];
    }

    checkPermissionsAndThrow(2, "body.setSelectedDataAsync");
    var commonParameters = parseCommonArgs(args, false, false);
    var parameters = {
      data: data
    };
    validateBodyParameters(parameters);
    addCoercionTypeParameter(parameters, commonParameters);

    if (parameters.coercionType === undefined) {
      invokeCallbackWithCoercionTypeError(commonParameters);
      return;
    }

    standardInvokeHostMethod(dispid, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
  };
}
// CONCATENATED MODULE: ./src/methods/setSignature.ts






function setSignature(data) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "item.body.setSignatureAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    data: data
  };

  if (isNullOrUndefined(parameters.data)) {
    parameters.data = "";
  } else {
    validateSignatureBodyParameters(parameters);
  }

  addCoercionTypeParameter(parameters, commonParameters);

  if (parameters.coercionType === undefined) {
    invokeCallbackWithCoercionTypeError(commonParameters);
    return;
  }

  standardInvokeHostMethod(173, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}
// CONCATENATED MODULE: ./src/api/getBodySurface.ts








function getBodySurface(isCompose) {
  var body = objectDefine({}, {
    getAsync: getBody
  });

  if (isCompose) {
    objectDefine(body, {
      appendOnSendAsync: appendOnSend,
      getTypeAsync: getBodyType,
      prependAsync: bodyPrepend,
      setAsync: setBody,
      setSelectedDataAsync: setSelectedData(13),
      setSignatureAsync: setSignature
    });
  }

  return body;
}
// CONCATENATED MODULE: ./src/methods/getAllInternetHeaders.ts



function getAllInternetHeaders() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "item.getAllInternetHeadersAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  standardInvokeHostMethod(168, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/types/ItemNotificationMessageType.ts
var ItemNotificationMessageType;

(function (ItemNotificationMessageType) {
  ItemNotificationMessageType[ItemNotificationMessageType["informationalMessage"] = 0] = "informationalMessage";
  ItemNotificationMessageType[ItemNotificationMessageType["progressIndicator"] = 1] = "progressIndicator";
  ItemNotificationMessageType[ItemNotificationMessageType["errorMessage"] = 2] = "errorMessage";
  ItemNotificationMessageType[ItemNotificationMessageType["insightMessage"] = 3] = "insightMessage";
})(ItemNotificationMessageType || (ItemNotificationMessageType = {}));
// CONCATENATED MODULE: ./src/utils/validateString.ts


function validateStringParam(paramName, paramValue) {
  if (isNullOrUndefined(paramValue) || paramValue === "") {
    throw createNullArgumentError(paramName);
  }

  if (!(typeof paramValue === "string")) {
    throw createArgumentTypeError(paramName, typeof paramValue, "string");
  }
}
// CONCATENATED MODULE: ./src/validation/notificationMessagesConstants.ts
var MaximumKeyLength = 32;
var MaximumIconLength = 32;
var MaximumMessageLength = 150;
var MaximumActionTextLength = 30;
var NotificationsKeyParameterName = "key";
var NotificationsTypeParameterName = "type";
var NotificationsIconParameterName = "icon";
var NotificationsMessageParameterName = "message";
var NotificationsPersistentParameterName = "persistent";
var NotificationsActionsDefinitionParameterName = "actions";
var NotificationsActionTypeParameterName = "actionType";
var NotificationsActionTextParameterName = "actionText";
var NotificationsActionCommandIdParameterName = "commandId";
var NotificationsActionShowTaskPaneActionId = "showTaskPane";
// CONCATENATED MODULE: ./src/validation/validateNotificationMessages.ts






function validateKey(key) {
  validateStringParam(NotificationsKeyParameterName, key);

  if (key.length > MaximumKeyLength) {
    throw createArgumentOutOfRange(NotificationsKeyParameterName, key.length);
  }
}
function validateData(data) {
  validateStringParam(NotificationsTypeParameterName, data.type);

  if (data.type === MailboxEnums.ItemNotificationMessageType.InformationalMessage) {
    validateStringParam(NotificationsIconParameterName, data.icon);

    if (data.icon.length > MaximumIconLength) {
      throw createArgumentOutOfRange(NotificationsIconParameterName, data.icon.length);
    }

    if (isNullOrUndefined(data.persistent)) {
      throw createNullArgumentError(NotificationsPersistentParameterName);
    }

    if (typeof data.persistent !== "boolean") {
      throw createArgumentTypeError(NotificationsPersistentParameterName, typeof data.persistent, "boolean");
    }

    if (!isNullOrUndefined(data.actions)) {
      throw createArgumentError(NotificationsActionsDefinitionParameterName, getString("l_ActionsDefinitionWrongNotificationMessageError_Text"));
    }
  } else if (data.type === MailboxEnums.ItemNotificationMessageType.InsightMessage) {
    validateInsightMessageParameters(data);
  } else {
    if (!isNullOrUndefined(data.icon)) {
      throw createArgumentError(NotificationsIconParameterName);
    }

    if (!isNullOrUndefined(data.persistent)) {
      throw createArgumentError(NotificationsPersistentParameterName);
    }

    if (!isNullOrUndefined(data.actions)) {
      throw createArgumentError(NotificationsActionsDefinitionParameterName, getString("l_ActionsDefinitionWrongNotificationMessageError_Text"));
    }
  }

  validateStringParam(NotificationsMessageParameterName, data.message);

  if (data.message.length > MaximumMessageLength) {
    throw createArgumentOutOfRange(NotificationsMessageParameterName, data.message.length);
  }
}

function validateInsightMessageParameters(data) {
  validateStringParam(NotificationsIconParameterName, data.icon);

  if (data.icon.length > MaximumIconLength) {
    throw createArgumentOutOfRange(NotificationsIconParameterName, data.icon.length);
  }

  if (!isNullOrUndefined(data.persistent)) {
    throw createArgumentError(NotificationsPersistentParameterName);
  }

  if (isNullOrUndefined(data.actions)) {
    throw createNullArgumentError(NotificationsActionsDefinitionParameterName);
  } else {
    validateActionsDefinitionBlob(data.actions);
  }
}

function validateActionsDefinitionBlob(actionsDefinitionBlob) {
  var actionsDefinition = extractActionsDefinition(actionsDefinitionBlob);

  if (isNullOrUndefined(actionsDefinition)) {
    return;
  }

  validateActionsDefinitionActionsType(actionsDefinition);
  validateActionsDefinitionActionsText(actionsDefinition);
}

function extractActionsDefinition(actionsDefinitionBlob) {
  var actionsDefinition = null;

  if (Array.isArray(actionsDefinitionBlob)) {
    if (actionsDefinitionBlob.length === 1) {
      actionsDefinition = actionsDefinitionBlob[0];
    } else if (actionsDefinitionBlob.length > 1) {
      throw createArgumentError(NotificationsActionsDefinitionParameterName, getString("l_ActionsDefinitionMultipleActionsError_Text"));
    }
  } else {
    throw createArgumentError(NotificationsActionsDefinitionParameterName);
  }

  return actionsDefinition;
}

function validateActionsDefinitionActionsType(actionsDefinition) {
  if (isNullOrUndefined(actionsDefinition.actionType)) {
    throw createNullArgumentError(NotificationsActionTypeParameterName);
  }

  if (NotificationsActionShowTaskPaneActionId !== actionsDefinition.actionType) {
    throw createArgumentError(NotificationsActionTypeParameterName, getString("l_InvalidActionType_Text"));
  } else {
    if (isNullOrUndefined(actionsDefinition.commandId) || typeof actionsDefinition.commandId !== "string" || actionsDefinition.commandId === "") {
      throw createArgumentError(NotificationsActionCommandIdParameterName, getString("l_InvalidCommandIdError_Text"));
    }
  }
}

function validateActionsDefinitionActionsText(actionsDefinition) {
  if (isNullOrUndefined(actionsDefinition.actionText) || actionsDefinition.actionText === "" || typeof actionsDefinition.actionText !== "string") {
    throw createNullArgumentError(NotificationsActionTextParameterName);
  }

  if (actionsDefinition.actionText.length > MaximumActionTextLength) {
    throw createArgumentOutOfRange(NotificationsActionTextParameterName, actionsDefinition.actionText.length);
  }
}
// CONCATENATED MODULE: ./src/methods/addNotificationMessage.ts







function addNotificationMessage(key, data) {
  var args = [];

  for (var _i = 2; _i < arguments.length; _i++) {
    args[_i - 2] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "notificationMessages.addAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  validateKey(key);
  validateData(data);
  var type = ItemNotificationMessageType[data.type];

  if (isNullOrUndefined(type)) {
    throw createArgumentError("type");
  }

  var message = data.message;
  var icon = data.icon;
  var persistent = data.persistent;
  var actions = data.actions;
  var parameters = {
    key: key,
    message: message,
    type: type,
    icon: icon,
    persistent: persistent,
    actions: actions
  };
  standardInvokeHostMethod(33, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}
// CONCATENATED MODULE: ./src/methods/getAllNotificationMessages.ts



function getAllNotificationMessages() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "notificationMessages.getAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  standardInvokeHostMethod(34, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/methods/removeNotificationMessage.ts




function removeNotificationMessage(key) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "notificationMessages.removeAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  validateKey(key);
  var parameters = {
    key: key
  };
  standardInvokeHostMethod(36, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}
// CONCATENATED MODULE: ./src/methods/replaceNotificationMessage.ts







function replaceNotificationMessage(key, data) {
  var args = [];

  for (var _i = 2; _i < arguments.length; _i++) {
    args[_i - 2] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "notificationMessages.replaceAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  validateKey(key);
  validateData(data);
  var type = ItemNotificationMessageType[data.type];

  if (isNullOrUndefined(type)) {
    throw createArgumentError("type");
  }

  var message = data.message;
  var icon = data.icon;
  var persistent = data.persistent;
  var actions = data.actions;
  var parameters = {
    key: key,
    message: message,
    type: type,
    icon: icon,
    persistent: persistent,
    actions: actions
  };
  standardInvokeHostMethod(35, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}
// CONCATENATED MODULE: ./src/api/getNotificationMessagesSurface.ts





function getNotificationMessageSurface() {
  return objectDefine({}, {
    addAsync: addNotificationMessage,
    getAllAsync: getAllNotificationMessages,
    removeAsync: removeNotificationMessage,
    replaceAsync: replaceNotificationMessage
  });
}
// CONCATENATED MODULE: ./src/validation/validateDisplayReplyForm.ts





function validateStringParameters(formData) {
  if (!isNullOrUndefined(formData)) {
    throwOnOutOfRange(formData.length, 0, maxBodyLength, "htmlBody");
  }
}
function validateAndGetHtmlBody(data) {
  var htmlBody = "";

  if (data.htmlBody) {
    throwOnInvalidHtmlBody(data.htmlBody);
    htmlBody = data.htmlBody;
  }

  return htmlBody;
}
function throwOnInvalidHtmlBody(htmlBody) {
  if (!(typeof htmlBody === "string")) {
    throw createArgumentTypeError("htmlBody", typeof htmlBody, "string");
  }

  if (isNullOrUndefined(htmlBody)) {
    throw createNullArgumentError("htmlBody");
  }

  throwOnOutOfRange(htmlBody.length, 0, maxBodyLength, "htmlBody");
}
function validateAndGetAttachments(data) {
  var attachments = [];

  if (data.attachments) {
    attachments = data.attachments;
    throwOnInvalidAttachmentsArray(attachments);
  }

  return attachments;
}
// CONCATENATED MODULE: ./src/utils/getOptionsAndCallback.ts

function getOptionsAndCallback(data) {
  var args = [];

  if (!isNullOrUndefined(data.options)) {
    args[0] = data.options;
  }

  if (!isNullOrUndefined(data.callback)) {
    args[args.length] = data.callback;
  }

  return args;
}
// CONCATENATED MODULE: ./src/methods/displayReplyForm.ts
var displayReplyForm_spreadArrays = undefined && undefined.__spreadArrays || function () {
  for (var s = 0, i = 0, il = arguments.length; i < il; i++) {
    s += arguments[i].length;
  }

  for (var r = Array(s), k = 0, i = 0; i < il; i++) {
    for (var a = arguments[i], j = 0, jl = a.length; j < jl; j++, k++) {
      r[k] = a[j];
    }
  }

  return r;
};










function displayReplyForm(formData) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  return displayReplyFormHelper.apply(void 0, displayReplyForm_spreadArrays([false, false, formData], args));
}
function displayReplyAllForm(formData) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  return displayReplyFormHelper.apply(void 0, displayReplyForm_spreadArrays([true, false, formData], args));
}
function displayReplyFormAsync(formData) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkFeatureEnabledAndThrow(Features.displayXAsync, "displayReplyFormAsync");
  return displayReplyFormHelper.apply(void 0, displayReplyForm_spreadArrays([false, true, formData], args));
}
function displayReplyAllFormAsync(formData) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkFeatureEnabledAndThrow(Features.displayXAsync, "displayReplyAllFormAsync");
  return displayReplyFormHelper.apply(void 0, displayReplyForm_spreadArrays([true, true, formData], args));
}

function displayReplyFormHelper(isReplyAll, isAsync, formData) {
  var args = [];

  for (var _i = 3; _i < arguments.length; _i++) {
    args[_i - 3] = arguments[_i];
  }

  var dispidToInvoke;
  checkPermissionsAndThrow(1, "mailbox.displayReplyForm");
  var commonParameters = parseCommonArgs(getOptionsAndCallback(formData), false, false);

  if (isFeatureEnabled(Features.replyCallback)) {
    if (isNullOrUndefined(commonParameters) || commonParameters.options === undefined && commonParameters.callback === undefined) {
      commonParameters = parseCommonArgs(args, false, false);
    }
  }

  var parameters = {
    formData: formData
  };
  var updatedHtmlBody = null;
  var updateAttachments = null;

  if (typeof parameters.formData === "string") {
    if (isReplyAll) {
      if (isAsync) {
        dispidToInvoke = 184;
      } else {
        dispidToInvoke = 11;
      }
    } else {
      if (isAsync) {
        dispidToInvoke = 183;
      } else {
        dispidToInvoke = 10;
      }
    }

    validateStringParameters(parameters.formData);
    standardInvokeHostMethod(dispidToInvoke, commonParameters.asyncContext, commonParameters.callback, {
      htmlBody: parameters.formData
    }, undefined);
  } else if (typeof parameters.formData === "object") {
    updatedHtmlBody = validateAndGetHtmlBody(parameters.formData);
    updateAttachments = createAttachmentsDataForHost(validateAndGetAttachments(parameters.formData));

    if (isReplyAll) {
      if (isAsync) {
        dispidToInvoke = 182;
      } else {
        dispidToInvoke = 31;
      }
    } else {
      if (isAsync) {
        dispidToInvoke = 181;
      } else {
        dispidToInvoke = 30;
      }
    }

    standardInvokeHostMethod(dispidToInvoke, commonParameters.asyncContext, commonParameters.callback, {
      htmlBody: updatedHtmlBody,
      attachments: updateAttachments
    }, undefined);
  } else {
    throw createArgumentError();
  }
}
// CONCATENATED MODULE: ./src/methods/addCategories.ts




function addCategories(categories) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "categories.addAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    categories: categories
  };
  validateCategoryArray(categories);
  standardInvokeHostMethod(158, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}
// CONCATENATED MODULE: ./src/methods/getCategories.ts



function getCategories() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "categories.getAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  standardInvokeHostMethod(157, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/methods/removeCategories.ts




function removeCategories(categories) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "categories.removeAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    categories: categories
  };
  validateCategoryArray(categories);
  standardInvokeHostMethod(159, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}
// CONCATENATED MODULE: ./src/api/getCategoriesSurface.ts




function getCategoriesSurface() {
  return objectDefine({}, {
    addAsync: addCategories,
    getAsync: getCategories,
    removeAsync: removeCategories
  });
}
// CONCATENATED MODULE: ./src/methods/getAttachmentContent.ts




function getAttachmentContent(id) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "item.getAttachmentContentAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  var parameters = {
    id: id
  };
  getAttachmentContent_validateParameters(parameters);
  standardInvokeHostMethod(150, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}

function getAttachmentContent_validateParameters(parameters) {
  validateStringParam("attachmentId", parameters.id);
}
// CONCATENATED MODULE: ./src/methods/moveToFolder.ts





var Folder = MailboxEnums.Folder;
function moveToFolder(destinationFolder) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(3, "item.move");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    destinationFolder: destinationFolder
  };
  moveToFolder_validateParameters(destinationFolder);
  standardInvokeHostMethod(101, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}

function moveToFolder_validateParameters(destinationFolder) {
  if (destinationFolder !== Folder.Inbox && destinationFolder !== Folder.Junk && destinationFolder !== Folder.DeletedItems) {
    throw createArgumentError("destinationFolder");
  }
}
// CONCATENATED MODULE: ./src/utils/createEmailAddressDetails.ts

var ResponseType = MailboxEnums.ResponseType;
var RecipientType = MailboxEnums.RecipientType;
var responseTypeMap = [ResponseType.None, ResponseType.Organizer, ResponseType.Tentative, ResponseType.Accepted, ResponseType.Declined];
var recipientTypeMap = [RecipientType.Other, RecipientType.DistributionList, RecipientType.User, RecipientType.ExternalUser];
var createEmailAddressDetails = function createEmailAddressDetails(input) {
  var response = input.appointmentResponse;
  var type = input.recipientType;
  var emailAddressDetails = {
    emailAddress: input.address,
    displayName: input.name
  };

  if (typeof input.appointmentResponse === "number") {
    emailAddressDetails.appointmentResponse = response < responseTypeMap.length ? responseTypeMap[response] : ResponseType.None;
  }

  if (typeof input.recipientType === "number") {
    emailAddressDetails.recipientType = type < recipientTypeMap.length ? recipientTypeMap[type] : RecipientType.Other;
  }

  return emailAddressDetails;
};
function createEmailAddressDetailsForEntity(data) {
  return createEmailAddressDetails({
    name: data.Name || "",
    address: data.UserId || ""
  });
}
// CONCATENATED MODULE: ./src/methods/getDelayDelivery.ts



function getDelayDelivery() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "delayDeliveryTime.getAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  standardInvokeHostMethod(166, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/methods/setDelayDelivery.ts






function setDelayDelivery(dateTime) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "delayDeliveryTime.setAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  validateParamerters(dateTime);
  standardInvokeHostMethod(167, commonParameters.asyncContext, commonParameters.callback, {
    time: dateTime.getTime()
  }, undefined);
}

function validateParamerters(dateTime) {
  if (isNullOrUndefined(dateTime)) {
    throw createNullArgumentError("dateTime", "You cannot conduct to a null dateTime");
  }

  if (!(dateTime instanceof Date)) {
    throw createArgumentTypeError("dateTime", typeof dateTime, typeof Date);
  }

  if (isNaN(dateTime.getTime())) {
    throw createArgumentError("dateTime");
  }

  throwOnOutOfRange(dateTime.getTime(), -8640000000000000, 8640000000000000, "dateTime");
}
// CONCATENATED MODULE: ./src/api/getDelayDeliverySurface.ts



function getDelayDeliverySurface(isCompose) {
  var delayDelivery = objectDefine({}, {
    getAsync: getDelayDelivery
  });

  if (isCompose) {
    objectDefine(delayDelivery, {
      setAsync: setDelayDelivery
    });
  }

  return delayDelivery;
}
// CONCATENATED MODULE: ./src/utils/removeDuplicates.ts
function removeDuplicates(array, comparator) {
  for (var matchIndex1 = array.length - 1; matchIndex1 >= 0; matchIndex1--) {
    var removeMatch = false;

    for (var matchIndex2 = matchIndex1 - 1; matchIndex2 >= 0; matchIndex2--) {
      if (comparator(array[matchIndex1], array[matchIndex2])) {
        removeMatch = true;
        break;
      }
    }

    if (removeMatch) {
      array.splice(matchIndex1, 1);
    }
  }

  return array;
}
var stringComparator = function stringComparator(a, b) {
  return a === b;
};
var meetingComparator = function meetingComparator(a, b) {
  if (a === b) {
    return true;
  } else if (!a || !b) {
    return false;
  } else {
    return a.meetingString === b.meetingString;
  }
};
var taskComparator = function taskComparator(a, b) {
  if (a === b) {
    return true;
  } else if (!a || !b) {
    return false;
  } else {
    return a.taskString === b.taskString;
  }
};
var contactComparator = function contactComparator(a, b) {
  if (a === b) {
    return true;
  } else if (!a || !b) {
    return false;
  } else {
    return a.contactString === b.contactString;
  }
};
// CONCATENATED MODULE: ./src/utils/isLegacyEntityExtraction.ts

function isLegacyEntityExtraction() {
  return !!getInitialDataProp("entities") && getInitialDataProp("entities").IsLegacyExtraction !== undefined && getInitialDataProp("entities").IsLegacyExtraction;
}
// CONCATENATED MODULE: ./src/utils/resolveDate.ts


var totalBits = 18;
var typeBits = 3;
var preciseDateTypeBits = 3;
var yearBits = 7;
var monthBits = 4;
var dayBits = 5;
var modifierBits = 2;
var unitBits = 3;
var offsetBits = 6;
var tagBits = 4;
var preciseDateType = 0;
var relativeDateType = 1;
var oneDayInMilliseconds = 86400000;
var baseDate = new Date("0001-01-01T00:00:00Z");
function resolveDate(input, sentTime) {
  if (!sentTime) {
    return input;
  }

  var date = null;

  try {
    var sentDate = new Date(sentTime.getFullYear(), sentTime.getMonth(), sentTime.getDate(), 0, 0, 0, 0);
    var extractedDate = decode(input);

    if (!extractedDate) {
      return input;
    } else {
      var preciseDate = extractedDate;

      if (preciseDate.day && preciseDate.month && preciseDate.year !== undefined) {
        date = resolvePreciseDate(sentDate, extractedDate);
      } else {
        var relativeDate = extractedDate;

        if (relativeDate.modifier !== undefined && relativeDate.offset !== undefined && relativeDate.tag !== undefined && relativeDate.unit !== undefined) {
          date = resolveRelativeDate(sentDate, extractedDate);
        } else {
          date = sentDate;
        }
      }
    }

    if (Number.isNaN(date.getTime())) {
      return sentTime;
    }

    date.setMilliseconds(date.getMilliseconds() + (isLegacyEntityExtraction() ? getTimeOfDayInMillisecondsUTC(input) : getTimeOfDayInMilliseconds(input)));
    return date;
  } catch (e) {
    return sentTime;
  }
}

function decode(input) {
  var dateValueMask = (1 << totalBits - typeBits) - 1;
  var time = 0;

  if (input == null) {
    return undefined;
  }

  if (isLegacyEntityExtraction()) {
    time = getTimeOfDayInMillisecondsUTC(input);
  } else {
    time = getTimeOfDayInMilliseconds(input);
  }

  var inDateAtMidnight = input.getTime() - time;
  var value = (inDateAtMidnight - baseDate.getTime()) / oneDayInMilliseconds;

  if (value < 0) {
    return undefined;
  } else if (value >= 1 << totalBits) {
    return undefined;
  } else {
    var type = value >> totalBits - typeBits;
    value = value & dateValueMask;

    switch (type) {
      case preciseDateType:
        return decodePreciseDate(value);

      case relativeDateType:
        return decodeRelativeDate(value);

      default:
        return undefined;
    }
  }
}

function decodePreciseDate(value) {
  var cSubTypeMask = (1 << preciseDateTypeBits) - 1;
  var cMonthMask = (1 << monthBits) - 1;
  var cDayMask = (1 << dayBits) - 1;
  var cYearMask = (1 << yearBits) - 1;
  var year = 0;
  var month = 0;
  var day = 0;
  var subType = value >> totalBits - typeBits - preciseDateTypeBits & cSubTypeMask;

  if ((subType & 4) == 4) {
    year = value >> totalBits - typeBits - preciseDateTypeBits - yearBits & cYearMask;

    if ((subType & 2) == 2) {
      if ((subType & 1) == 1) {
        return undefined;
      }

      month = value >> totalBits - typeBits - preciseDateTypeBits - yearBits - monthBits & cMonthMask;
    }
  } else {
    if ((subType & 2) == 2) {
      month = value >> totalBits - typeBits - preciseDateTypeBits - monthBits & cMonthMask;
    }

    if ((subType & 1) == 1) {
      day = value >> totalBits - typeBits - preciseDateTypeBits - monthBits - dayBits & cDayMask;
    }
  }

  return createPreciseDate(day, month, year);
}

function resolvePreciseDate(sentDate, precise) {
  var year = precise.year;
  var month = precise.month == 0 ? sentDate.getMonth() : precise.month - 1;
  var day = precise.day;

  if (day == 0) {
    return sentDate;
  }

  var candidate;

  if (isNullOrUndefined(year)) {
    candidate = new Date(sentDate.getFullYear(), month, day);

    if (candidate.getTime() < sentDate.getTime()) {
      candidate = new Date(sentDate.getFullYear() + 1, month, day);
    }
  } else {
    candidate = new Date(year < 50 ? 2000 + year : 1900 + year, month, day);
  }

  if (candidate.getMonth() != month) {
    return sentDate;
  }

  return candidate;
}

function resolveRelativeDate(sentDate, relative) {
  var date;

  switch (relative.unit) {
    case 0:
      date = new Date(sentDate.getFullYear(), sentDate.getMonth(), sentDate.getDate());
      date.setDate(date.getDate() + relative.offset);
      return date;

    case 5:
      return findBestDateForWeekDate(sentDate, relative.offset, relative.tag);

    case 2:
      {
        var days = 1;

        switch (relative.modifier) {
          case 1:
            break;

          case 2:
            days = 16;
            break;

          default:
            if (relative.offset == 0) {
              days = sentDate.getDate();
            }

            break;
        }

        date = new Date(sentDate.getFullYear(), sentDate.getMonth(), days);
        date.setMonth(date.getMonth() + relative.offset);

        if (date.getTime() < sentDate.getTime()) {
          date.setDate(date.getDate() + sentDate.getDate() - 1);
        }

        return date;
      }

    case 1:
      date = new Date(sentDate.getFullYear(), sentDate.getMonth(), sentDate.getDate());
      date.setDate(sentDate.getDate() + 7 * relative.offset);

      if (relative.modifier == 1 || relative.modifier == 0) {
        date.setDate(date.getDate() + 1 - date.getDay());

        if (date.getTime() < sentDate.getTime()) {
          return sentDate;
        }

        return date;
      } else if (relative.modifier == 2) {
        date.setDate(date.getDate() + 5 - date.getDay());
        return date;
      }

      break;

    case 4:
      return findBestDateForWeekOfMonthDate(sentDate, relative);

    case 3:
      if (relative.offset > 0) {
        return new Date(sentDate.getFullYear() + relative.offset, 0, 1);
      }

      break;

    default:
      break;
  }

  return sentDate;
}

function findBestDateForWeekDate(sentDate, offset, tag) {
  if (offset > -5 && offset < 5) {
    var dayOfWeek = (tag + 6) % 7 + 1;
    var days = 7 * offset + (dayOfWeek - sentDate.getDay());
    sentDate.setDate(sentDate.getDate() + days);
    return sentDate;
  } else {
    var days = (tag - sentDate.getDay()) % 7;

    if (days < 0) {
      days += 7;
    }

    sentDate.setDate(sentDate.getDate() + days);
    return sentDate;
  }
}

function findBestDateForWeekOfMonthDate(sentDate, relative) {
  var date;
  var firstDay;
  var newDate;
  date = sentDate;

  if (relative.tag <= 0 || relative.tag > 12 || relative.offset <= 0 || relative.offset > 5) {
    return sentDate;
  }

  var monthOffset = (12 + relative.tag - date.getMonth() - 1) % 12;
  firstDay = new Date(date.getFullYear(), date.getMonth() + monthOffset, 1);

  if (relative.modifier == 1) {
    if (relative.offset == 1 && firstDay.getDay() != 6 && firstDay.getDay() != 0) {
      return firstDay;
    } else {
      newDate = new Date(firstDay.getFullYear(), firstDay.getMonth(), firstDay.getDate());
      newDate.setDate(newDate.getDate() + (7 + (1 - firstDay.getDay())) % 7);

      if (firstDay.getDay() != 6 && firstDay.getDay() != 0 && firstDay.getDay() != 1) {
        newDate.setDate(newDate.getDate() - 7);
      }

      newDate.setDate(newDate.getDate() + 7 * (relative.offset - 1));

      if (newDate.getMonth() + 1 != relative.tag) {
        return sentDate;
      }

      return newDate;
    }
  } else {
    newDate = new Date(firstDay.getFullYear(), firstDay.getMonth(), daysInMonth(firstDay.getMonth(), firstDay.getFullYear()));
    var offset = 1 - newDate.getDay();

    if (offset > 0) {
      offset = offset - 7;
    }

    newDate.setDate(newDate.getDate() + offset);
    newDate.setDate(newDate.getDate() + 7 * (1 - relative.offset));

    if (newDate.getMonth() + 1 != relative.tag) {
      if (firstDay.getDay() != 6 && firstDay.getDay() != 0) {
        return firstDay;
      } else {
        return sentDate;
      }
    } else {
      return newDate;
    }
  }
}

function decodeRelativeDate(value) {
  var tagMask = (1 << tagBits) - 1;
  var offsetMask = (1 << offsetBits) - 1;
  var unitMask = (1 << unitBits) - 1;
  var modifierMask = (1 << modifierBits) - 1;
  var tag = value & tagMask;
  value >>= tagBits;
  var offset = fromComplement(value & offsetMask, offsetBits);
  value >>= offsetBits;
  var unit = value & unitMask;
  value >>= unitBits;
  var modifier = value & modifierMask;

  try {
    return createRelativeDate(modifier, offset, unit, tag);
  } catch (_a) {
    return undefined;
  }
}

function fromComplement(value, n) {
  var signed = 1 << n - 1;
  var mask = (1 << n) - 1;

  if ((value & signed) == signed) {
    return -((value ^ mask) + 1);
  } else {
    return value;
  }
}

function daysInMonth(month, year) {
  return 32 - new Date(year, month, 32).getDate();
}

function getTimeOfDayInMilliseconds(inputTime) {
  var timeOfDay = 0;
  timeOfDay += inputTime.getHours() * 3600;
  timeOfDay += inputTime.getMinutes() * 60;
  timeOfDay += inputTime.getSeconds();
  timeOfDay *= 1000;
  timeOfDay += inputTime.getMilliseconds();
  return timeOfDay;
}

function getTimeOfDayInMillisecondsUTC(inputTime) {
  var timeOfDay = 0;
  timeOfDay += inputTime.getUTCHours() * 3600;
  timeOfDay += inputTime.getUTCMinutes() * 60;
  timeOfDay += inputTime.getUTCSeconds();
  timeOfDay *= 1000;
  timeOfDay += inputTime.getUTCMilliseconds();
  return timeOfDay;
}

function createPreciseDate(day, month, year) {
  return {
    day: day,
    month: month,
    year: year % 100
  };
}

function createRelativeDate(modifier, offset, unit, tag) {
  return {
    modifier: modifier,
    offset: offset,
    unit: unit,
    tag: tag
  };
}
// CONCATENATED MODULE: ./src/utils/findOffset.ts



function findOffset(value) {
  var ranges = getInitialDataProp("timeZoneOffsets");

  for (var r = 0; r < ranges.length; r++) {
    var range = ranges[r];
    var start = parseInt(range.start);
    var end = parseInt(range.end);

    if (value.getTime() - start >= 0 && value.getTime() - end < 0) {
      return parseInt(range.offset);
    }
  }

  throw createArgumentError("input", getString("l_InvalidDate_Text"));
}
// CONCATENATED MODULE: ./src/utils/convertToUtcClientTime.ts





function convertToUtcClientTime(input) {
  var retValue = localClientTimeToDate(input);

  if (!isNullOrUndefined(getInitialDataProp("timeZoneOffsets"))) {
    var offset = findOffset(retValue);
    retValue.setUTCMinutes(retValue.getUTCMinutes() - offset);
    offset = !input["timezoneOffset"] ? retValue.getTimezoneOffset() * -1 : input["timezoneOffset"];
    retValue.setUTCMinutes(retValue.getUTCMinutes() + offset);
  }

  return retValue;
}
function localClientTimeToDate(input) {
  var retValue = new Date(input["year"], input["month"], input["date"], input["hours"], input["minutes"], input["seconds"], input["milliseconds"] === null ? 0 : input["milliseconds"]);

  if (isNaN(retValue.getTime())) {
    throw createArgumentError("input", getString("l_InvalidDate_Text"));
  }

  return retValue;
}
// CONCATENATED MODULE: ./src/utils/dateToDictionary.ts
function dateToDictionary(date) {
  return {
    month: date.getMonth(),
    date: date.getDate(),
    year: date.getFullYear(),
    hours: date.getHours(),
    minutes: date.getMinutes(),
    seconds: date.getSeconds(),
    milliseconds: date.getMilliseconds()
  };
}
// CONCATENATED MODULE: ./src/utils/createEntities.ts









var EntityKeys;

(function (EntityKeys) {
  EntityKeys["meetingSuggestion"] = "MeetingSuggestions";
  EntityKeys["taskSuggestion"] = "TaskSuggestions";
  EntityKeys["address"] = "Addresses";
  EntityKeys["emailAddress"] = "EmailAddresses";
  EntityKeys["url"] = "Urls";
  EntityKeys["phoneNumber"] = "PhoneNumbers";
  EntityKeys["contact"] = "Contacts";
  EntityKeys["flightReservations"] = "FlightReservations";
  EntityKeys["parcelDeliveries"] = "ParcelDeliveries";
})(EntityKeys || (EntityKeys = {}));

function createEntities(data) {
  if (isNullOrUndefined(data)) {
    return {
      addresses: [],
      emailAddresses: [],
      urls: [],
      taskSuggestions: [],
      meetingSuggestions: [],
      phoneNumbers: [],
      contacts: [],
      flightReservations: [],
      parcelDelivery: []
    };
  } else {
    return {
      addresses: createEntities_createAddresses(data[EntityKeys.address]),
      emailAddresses: createEntities_createEmailAddresses(data[EntityKeys.emailAddress]),
      urls: createUrls(data[EntityKeys.url]),
      taskSuggestions: createEntities_createTaskSuggestions(data[EntityKeys.taskSuggestion]),
      meetingSuggestions: createEntities_createMeetingSuggestions(data[EntityKeys.meetingSuggestion]),
      phoneNumbers: createPhoneNumbers(data[EntityKeys.phoneNumber]),
      contacts: createEntities_createContacts(data[EntityKeys.contact]),
      flightReservations: createEntities_createReadItemArray(data[EntityKeys.flightReservations]),
      parcelDelivery: createEntities_createReadItemArray(data[EntityKeys.parcelDeliveries])
    };
  }
}
function createFilteredEntities(data, name) {
  checkPermissionsAndThrow(1, "item.getFilteredEntitiesByName");
  var results = Object.keys(data).map(function (entities) {
    var results = data[entities][name];
    if (results) return {
      entityType: entities,
      name: name,
      entities: data[entities][name]
    };else return undefined;
  }).filter(function (results) {
    return results !== undefined;
  });

  if (results.length === 0) {
    return null;
  }

  var matchedRule = results[0];

  switch (matchedRule.entityType) {
    case EntityKeys.meetingSuggestion:
      return createEntities_createMeetingSuggestions(matchedRule.entities);

    case EntityKeys.address:
      return createEntities_createAddresses(matchedRule.entities);

    case EntityKeys.contact:
      return createEntities_createContacts(matchedRule.entities);

    case EntityKeys.emailAddress:
      return createEntities_createEmailAddresses(matchedRule.entities);

    case EntityKeys.phoneNumber:
      return createPhoneNumbers(matchedRule.entities);

    case EntityKeys.taskSuggestion:
      return createEntities_createTaskSuggestions(matchedRule.entities);

    case EntityKeys.url:
      return createUrls(matchedRule.entities);

    default:
      return createEntities_createReadItemArray(matchedRule.entities);
  }
}
var createEntities_createAddresses = function createAddresses(data) {
  var addresses = data || [];
  return removeDuplicates(addresses, stringComparator);
};
var createEntities_createEmailAddresses = function createEmailAddresses(data) {
  if (getPermissionLevel_getPermissionLevel() === 0) {
    return [];
  }

  return data || [];
};
var createUrls = function createUrls(data) {
  return data || [];
};
var createEntities_createTaskSuggestions = function createTaskSuggestions(data) {
  if (getPermissionLevel_getPermissionLevel() === 0) {
    return [];
  }

  var tasks = data || [];
  tasks = tasks.map(function (task) {
    return {
      assignees: (task.Assignees || []).map(createEmailAddressDetailsForEntity),
      taskString: task.TaskString
    };
  });
  return removeDuplicates(tasks, taskComparator);
};
var createEntities_createMeetingSuggestions = function createMeetingSuggestions(data) {
  if (getPermissionLevel_getPermissionLevel() === 0) {
    return [];
  }

  var meetings = data || [];
  meetings = meetings.map(function (meeting) {
    var start = getDate(meeting.StartTime);
    var end = getDate(meeting.EndTime);
    return {
      meetingString: meeting.MeetingString,
      attendees: (meeting.Attendees || []).map(createEmailAddressDetailsForEntity),
      location: meeting.Location,
      subject: meeting.Subject,
      start: meeting.StartTime !== undefined ? start : undefined,
      end: meeting.EndTime !== undefined ? end : undefined
    };
  });
  return removeDuplicates(meetings, meetingComparator);
};

function getDate(date) {
  var result = resolveDate(new Date(date), new Date(getInitialDataProp("dateTimeSent")));

  if (result.getTime() !== new Date(date).getTime()) {
    return convertToUtcClientTime(dateToDictionary(result));
  }

  return new Date(date);
}

var createPhoneNumbers = function createPhoneNumbers(data) {
  var phoneNumbers = data || [];
  return phoneNumbers.map(function (number) {
    return {
      phoneString: number.PhoneString,
      originalPhoneString: number.OriginalPhoneString,
      type: number.Type
    };
  });
};
var createEntities_createContacts = function createContacts(data) {
  if (getPermissionLevel_getPermissionLevel() === 0) {
    return [];
  }

  var contacts = data || [];
  contacts = contacts.map(function (contact) {
    return {
      personName: contact.PersonName,
      businessName: contact.BusinessName,
      phoneNumbers: createPhoneNumbers(contact.PhoneNumbers || []),
      emailAddresses: contact.EmailAddresses || [],
      urls: contact.Urls || [],
      addresses: contact.Addresses || [],
      contactString: contact.ContactString
    };
  });
  return removeDuplicates(contacts, contactComparator);
};
var createEntities_createReadItemArray = function createReadItemArray(data) {
  if (getPermissionLevel_getPermissionLevel() === 0) {
    return [];
  }

  return data || [];
};
// CONCATENATED MODULE: ./src/api/Entities.ts



var entityPermissions = {
  meetingSuggestion: 1,
  taskSuggestion: 1,
  address: 0,
  emailAddress: 1,
  url: 0,
  phoneNumber: 0,
  contact: 1,
  flightReservations: 1,
  parcelDeliveries: 1
};
var entityKeys = {
  meetingSuggestion: "meetingSuggestions",
  taskSuggestion: "taskSuggestions",
  address: "addresses",
  emailAddress: "emailAddresses",
  url: "urls",
  phoneNumber: "phoneNumbers",
  contact: "contacts",
  flightReservations: "flightReservations",
  parcelDeliveries: "parcelDeliveries"
};
var Entities_getEntities = function getEntities() {
  return createEntities(getInitialDataProp("entities"));
};
var Entities_getEntitiesByType = function getEntitiesByType(entityType) {
  var entities = createEntities(getInitialDataProp("entities"));
  checkPermissionsAndThrow(entityPermissions[entityType] !== undefined ? entityPermissions[entityType] : 1, entityType);
  var entityProperty = entityKeys[entityType];

  if (entityProperty === undefined) {
    return null;
  }

  return entities[entityProperty];
};
var Entities_getFilteredEntitiesByName = function getFilteredEntitiesByName(name) {
  return createFilteredEntities(getInitialDataProp("filteredEntities"), name);
};
var Entities_getRegExMatches = function getRegExMatches() {
  return getInitialDataProp("regExMatches");
};
var Entities_getRegExMatchesByName = function getRegExMatchesByName(name) {
  var regExMatches = getInitialDataProp("regExMatches") || {};
  return regExMatches[name];
};
var Entities_getSelectedEntities = function getSelectedEntities() {
  return createEntities(getInitialDataProp("selectedEntities"));
};
var Entities_getSelectedRegExMatches = function getSelectedRegExMatches() {
  return getInitialDataProp("selectedRegExMatches");
};
// CONCATENATED MODULE: ./src/utils/CustomJsonAttachmentsResponse.ts


function CustomJsonAttachmentsResponse(arrayOfAttachmentJsonData) {
  var customJsonResponse = [];

  if (getPermissionLevel_getPermissionLevel() === 0) {
    return [];
  }

  if (!!arrayOfAttachmentJsonData) {
    for (var i = 0; i < arrayOfAttachmentJsonData.length; i++) {
      if (!!arrayOfAttachmentJsonData[i]) {
        var newAttachment = convertAttachmentType(arrayOfAttachmentJsonData[i]);
        customJsonResponse.push(newAttachment);
      }
    }
  }

  return customJsonResponse;
}
function convertAttachmentType(attachmentDetails) {
  if (attachmentDetails.attachmentType !== null || attachmentDetails.attachmentType !== undefined) {
    switch (attachmentDetails.attachmentType) {
      case 0:
        {
          attachmentDetails.attachmentType = MailboxEnums.AttachmentType.File;
          break;
        }

      case 1:
        {
          attachmentDetails.attachmentType = MailboxEnums.AttachmentType.Item;
          break;
        }

      case 2:
        {
          attachmentDetails.attachmentType = MailboxEnums.AttachmentType.Cloud;
          break;
        }
    }
  }

  return attachmentDetails;
}
// CONCATENATED MODULE: ./src/methods/deepClone.ts
function deepClone(original) {
  return JSON.parse(JSON.stringify(original));
}
// CONCATENATED MODULE: ./src/validation/seriesTimeConstants.ts
var StartYearKey = "startYear";
var StartMonthKey = "startMonth";
var StartDayKey = "startDay";
var EndYearKey = "endYear";
var EndMonthKey = "endMonth";
var EndDayKey = "endDay";
var NoEndDateKey = "noEndDate";
var StartTimeMinKey = "startTimeMin";
var DurationMinKey = "durationMin";
// CONCATENATED MODULE: ./src/validation/recurrenceConstants.ts
var StartDateKey = "startDate";
var EndDateKey = "endDate";
var StartTimeKey = "startTime";
var EndTimeKey = "endTime";
var RecurrenceTypeKey = "recurrenceType";
var SeriesTimeKey = "seriesTime";
var SeriesTimeJsonKey = "seriesTimeJson";
var RecurrenceTimeZoneKey = "recurrenceTimeZone";
var RecurrenceTimeZoneName = "name";
var RecurrencePropertiesKey = "recurrenceProperties";
var IntervalKey = "interval";
var DaysKey = "days";
var DayOfMonthKey = "dayOfMonth";
var DayOfWeekKey = "dayOfWeek";
var WeekNumberKey = "weekNumber";
var MonthKey = "month";
var FirstDayOfWeekKey = "firstDayOfWeek";
// CONCATENATED MODULE: ./src/utils/seriesTimeUtils.ts



function prependZeroToString(number) {
  if (number < 0) {
    number = 1;
  }

  if (number < 10) {
    return "0" + number.toString();
  }

  return number.toString();
}
function throwOnInvalidDate(year, month, day) {
  if (!isValidDate(year, month, day)) {
    throw createArgumentError(SeriesTimeKey, getString("l_InvalidDate_Text"));
  }
}
function isValidDate(year, month, day) {
  if (year < 1601 || month < 1 || month > 12 || day < 1 || day > 31) {
    return false;
  }

  return true;
}
function throwOnInvalidDateString(dateString) {
  var regEx = new RegExp("^\\d{4}-(?:[0]\\d|1[0-2])-(?:[0-2]\\d|3[01])$");

  if (!regEx.test(dateString)) {
    throw createArgumentError(SeriesTimeKey, getString("l_InvalidDate_Text"));
  }
}
// CONCATENATED MODULE: ./src/api/SeriesTime.ts






var SeriesTime_SeriesTime = function () {
  function SeriesTime() {
    this.startYear = 0;
    this.startMonth = 0;
    this.startDay = 0;
    this.endYear = 0;
    this.endMonth = 0;
    this.endDay = 0;
    this.startTimeMinutes = 0;
    this.durationMinutes = 0;
  }

  SeriesTime.prototype.getDuration = function () {
    return this.durationMinutes;
  };

  SeriesTime.prototype.getEndTime = function () {
    var endTimeMinutes = this.startTimeMinutes + this.durationMinutes;
    var minutes = endTimeMinutes % 60;
    var hours = Math.floor(endTimeMinutes / 60) % 24;
    return "T" + prependZeroToString(hours) + ":" + prependZeroToString(minutes) + ":00.000";
  };

  SeriesTime.prototype.getEndDate = function () {
    if (this.endYear === 0 && this.endMonth === 0 && this.endDay === 0) {
      return null;
    }

    return this.endYear.toString() + "-" + prependZeroToString(this.endMonth) + "-" + prependZeroToString(this.endDay);
  };

  SeriesTime.prototype.getStartDate = function () {
    return this.startYear.toString() + "-" + prependZeroToString(this.startMonth) + "-" + prependZeroToString(this.startDay);
  };

  SeriesTime.prototype.getStartTime = function () {
    var minutes = this.startTimeMinutes % 60;
    var hours = Math.floor(this.startTimeMinutes / 60);
    return "T" + prependZeroToString(hours) + ":" + prependZeroToString(minutes) + ":00.000";
  };

  SeriesTime.prototype.setDuration = function (minutes) {
    if (minutes >= 0) {
      this.durationMinutes = minutes;
    } else {
      throw createArgumentError(undefined, getString("l_InvalidTime_Text"));
    }
  };

  SeriesTime.prototype.setEndDate = function (yearOrDateString, month, day) {
    if (yearOrDateString !== null && !isNullOrUndefined(month) && day !== null) {
      this.setDateHelper(false, yearOrDateString, month, day);
    } else if (yearOrDateString !== null) {
      this.setDateHelper(false, yearOrDateString);
    } else if (yearOrDateString == null) {
      this.endYear = 0;
      this.endMonth = 0;
      this.endDay = 0;
    }
  };

  SeriesTime.prototype.setStartDate = function (yearOrDateString, month, day) {
    if (yearOrDateString !== null && !isNullOrUndefined(month) && day !== null) {
      this.setDateHelper(true, yearOrDateString, month, day);
    } else if (yearOrDateString !== null) {
      this.setDateHelper(true, yearOrDateString);
    }
  };

  SeriesTime.prototype.setStartTime = function (hoursOrTimeString, minutes) {
    if (!isNullOrUndefined(hoursOrTimeString) && !isNullOrUndefined(minutes)) {
      var totalMinutes = hoursOrTimeString * 60 + minutes;

      if (totalMinutes >= 0) {
        this.startTimeMinutes = totalMinutes;
      } else {
        throw createArgumentError(undefined, getString("l_InvalidTime_Text"));
      }
    } else if (!isNullOrUndefined(hoursOrTimeString)) {
      var timeString = hoursOrTimeString;
      var newDateString = "2017-01-15" + timeString + "Z";
      var regEx = new RegExp("^T[0-2]\\d:[0-5]\\d:[0-5]\\d\\.\\d{3}$");

      if (!regEx.test(timeString)) {
        throw createArgumentError(undefined, getString("l_InvalidTime_Text"));
      }

      var dateObject = new Date(newDateString);

      if (!isNullOrUndefined(dateObject) && !isNaN(dateObject.getUTCHours()) && !isNaN(dateObject.getUTCMinutes())) {
        this.startTimeMinutes = dateObject.getUTCHours() * 60 + dateObject.getUTCMinutes();
      } else {
        throw createArgumentError(undefined, getString("l_InvalidTime_Text"));
      }
    }
  };

  SeriesTime.prototype.isValid = function () {
    if (!isValidDate(this.startYear, this.startMonth, this.startDay)) {
      return false;
    }

    if (this.endDay !== 0 && this.endMonth !== 0 && this.endYear !== 0) {
      if (!isValidDate(this.endYear, this.endMonth, this.endDay)) {
        return false;
      }
    }

    if (this.startTimeMinutes < 0 || this.durationMinutes <= 0) {
      return false;
    }

    return true;
  };

  SeriesTime.prototype.exportToSeriesTimeJson = function () {
    var result = {};
    result[StartYearKey] = this.startYear;
    result[StartMonthKey] = this.startMonth;
    result[StartDayKey] = this.startDay;

    if (this.endYear === 0 && this.endMonth === 0 && this.endDay === 0) {
      result[NoEndDateKey] = true;
    } else {
      result[EndYearKey] = this.endYear;
      result[EndMonthKey] = this.endMonth;
      result[EndDayKey] = this.endDay;
    }

    result[StartTimeMinKey] = this.startTimeMinutes;

    if (this.durationMinutes > 0) {
      result[DurationMinKey] = this.durationMinutes;
    }

    return result;
  };

  SeriesTime.prototype.importFromSeriesTimeJsonObject = function (jsonObject) {
    this.startYear = jsonObject[StartYearKey];
    this.startMonth = jsonObject[StartMonthKey];
    this.startDay = jsonObject[StartDayKey];

    if (jsonObject[NoEndDateKey] != null && typeof jsonObject[NoEndDateKey] === "boolean") {
      this.endYear = 0;
      this.endMonth = 0;
      this.endDay = 0;
    } else {
      this.endYear = jsonObject[EndYearKey];
      this.endMonth = jsonObject[EndMonthKey];
      this.endDay = jsonObject[EndDayKey];
    }

    this.startTimeMinutes = jsonObject[StartTimeMinKey];
    this.durationMinutes = jsonObject[DurationMinKey];
  };

  SeriesTime.prototype.setDateHelper = function (isStart, yearOrDateString, month, day) {
    var yearCalculated = 0;
    var monthCalculated = 0;
    var dayCalculated = 0;

    if (yearOrDateString !== null && !isNullOrUndefined(month) && day !== null) {
      throwOnInvalidDate(yearOrDateString, month + 1, day);
      yearCalculated = yearOrDateString;
      monthCalculated = month + 1;
      dayCalculated = day;
    } else if (yearOrDateString !== null) {
      var dateString = yearOrDateString;
      throwOnInvalidDateString(dateString);
      var dateObject = new Date(dateString);

      if (dateObject !== null && !isNaN(dateObject.getUTCFullYear()) && !isNaN(dateObject.getUTCMonth()) && !isNaN(dateObject.getUTCDate())) {
        throwOnInvalidDate(dateObject.getUTCFullYear(), dateObject.getUTCMonth() + 1, dateObject.getUTCDate());
        yearCalculated = dateObject.getUTCFullYear();
        monthCalculated = dateObject.getUTCMonth() + 1;
        dayCalculated = dateObject.getUTCDate();
      }
    }

    if (yearCalculated !== 0 && monthCalculated !== 0 && dayCalculated !== 0) {
      if (isStart) {
        this.startYear = yearCalculated;
        this.startMonth = monthCalculated;
        this.startDay = dayCalculated;
      } else {
        this.endYear = yearCalculated;
        this.endMonth = monthCalculated;
        this.endDay = dayCalculated;
      }
    }
  };

  SeriesTime.prototype.isEndAfterStart = function () {
    if (this.endYear === 0 && this.endMonth === 0 && this.endDay === 0) {
      return true;
    }

    var startDateTime = new Date();
    startDateTime.setFullYear(this.startYear);
    startDateTime.setMonth(this.startMonth - 1);
    startDateTime.setDate(this.startDay);
    var endDateTime = new Date();
    endDateTime.setFullYear(this.endYear);
    endDateTime.setMonth(this.endMonth - 1);
    endDateTime.setDate(this.endDay);
    return endDateTime >= startDateTime;
  };

  return SeriesTime;
}();


// CONCATENATED MODULE: ./src/utils/recurrenceUtils.ts



function copyRecurrenceObjectConvertSeriesTimeJson(recurrenceOriginal) {
  if (isNullOrUndefined(recurrenceOriginal) || isNullOrUndefined(recurrenceOriginal.seriesTimeJson)) {
    return recurrenceOriginal;
  }

  var recurrenceCopy = {
    recurrenceType: "",
    recurrenceProperties: null,
    recurrenceTimeZone: null
  };
  var newSeriesTime = new SeriesTime_SeriesTime();

  if (!isNullOrUndefined(recurrenceOriginal.recurrenceProperties)) {
    recurrenceCopy.recurrenceProperties = deepClone(recurrenceOriginal.recurrenceProperties);
  }

  recurrenceCopy.recurrenceType = recurrenceOriginal.recurrenceType;

  if (!isNullOrUndefined(recurrenceOriginal.recurrenceTimeZone)) {
    recurrenceCopy.recurrenceTimeZone = deepClone(recurrenceOriginal.recurrenceTimeZone);
  }

  newSeriesTime.importFromSeriesTimeJsonObject(recurrenceOriginal.seriesTimeJson);
  recurrenceCopy.seriesTime = newSeriesTime;
  return recurrenceCopy;
}
// CONCATENATED MODULE: ./src/api/getMessageRead.ts
















function getMessageRead() {
  var sender = getInitialDataProp("sender");
  var from = getInitialDataProp("from");
  var dateTimeCreated = getInitialDataProp("dateTimeCreated");
  var dateTimeModified = getInitialDataProp("dateTimeModified");
  var end = getInitialDataProp("end");
  var start = getInitialDataProp("start");
  var messageRead = objectDefine({}, {
    attachments: CustomJsonAttachmentsResponse(getInitialDataProp("attachments")),
    bcc: (getInitialDataProp("bcc") || []).map(createEmailAddressDetails),
    body: getBodySurface(false),
    categories: getCategoriesSurface(),
    cc: (getInitialDataProp("cc") || []).map(createEmailAddressDetails),
    conversationId: getInitialDataProp("conversationId"),
    dateTimeCreated: dateTimeCreated ? new Date(dateTimeCreated) : undefined,
    dateTimeModified: dateTimeModified ? new Date(dateTimeModified) : undefined,
    end: end ? new Date(end) : undefined,
    from: from ? createEmailAddressDetails(from) : undefined,
    getAllInternetHeadersAsync: getAllInternetHeaders,
    internetMessageId: getInitialDataProp("internetMessageId"),
    itemClass: getInitialDataProp("itemClass"),
    itemId: getInitialDataProp("id"),
    itemType: "message",
    location: getInitialDataProp("location"),
    move: moveToFolder,
    normalizedSubject: getInitialDataProp("normalizedSubject"),
    notificationMessages: getNotificationMessageSurface(),
    recurrence: copyRecurrenceObjectConvertSeriesTimeJson(getInitialDataProp("recurrence")),
    seriesId: getInitialDataProp("seriesId"),
    sender: sender ? createEmailAddressDetails(sender) : undefined,
    start: start ? new Date(start) : undefined,
    subject: getInitialDataProp("subject"),
    to: (getInitialDataProp("to") || []).map(createEmailAddressDetails),
    displayReplyForm: displayReplyForm,
    displayReplyFormAsync: displayReplyFormAsync,
    displayReplyAllForm: displayReplyAllForm,
    displayReplyAllFormAsync: displayReplyAllFormAsync,
    getAttachmentContentAsync: getAttachmentContent,
    getEntities: Entities_getEntities,
    getEntitiesByType: Entities_getEntitiesByType,
    getFilteredEntitiesByName: Entities_getFilteredEntitiesByName,
    getInitializationContextAsync: getInitializationContext,
    getRegExMatches: Entities_getRegExMatches,
    getRegExMatchesByName: Entities_getRegExMatchesByName,
    getSelectedEntities: Entities_getSelectedEntities,
    getSelectedRegExMatches: Entities_getSelectedRegExMatches,
    loadCustomPropertiesAsync: loadCustomProperties,
    delayDeliveryTime: getDelayDeliverySurface(false),
    isAllDayEvent: getInitialDataProp("isAllDayEvent"),
    sensitivity: getInitialDataProp("sensitivity")
  });
  return messageRead;
}
// CONCATENATED MODULE: ./src/validation/validateAttachments.ts




function validateAddFileAttachmentApis(attachmentName) {
  if (isNullOrUndefined(attachmentName) || attachmentName === "" || !(typeof attachmentName === "string")) {
    throw createArgumentError("attachmentName");
  }

  throwOnOutOfRange(attachmentName.length, 0, MaxAttachmentNameLength, "attachmentName");
}
// CONCATENATED MODULE: ./src/validation/attachmentsConstants.ts
var AddItemAttachmentClientEndPointTimeoutInMilliseconds = 600000;
// CONCATENATED MODULE: ./src/methods/addFileAttachment.ts








function addFileAttachment(uri, attachmentName) {
  var args = [];

  for (var _i = 2; _i < arguments.length; _i++) {
    args[_i - 2] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "item.addFileAttachmentAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var isInline = false;

  if (!!commonParameters.options) {
    isInline = !!commonParameters.options.isInline;
  }

  var name = attachmentName;
  var parameters = {
    uri: uri,
    name: name,
    isInline: isInline,
    __timeout__: AddItemAttachmentClientEndPointTimeoutInMilliseconds
  };
  addFileAttachment_validateParameters(parameters);
  standardInvokeHostMethod(16, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}

function addFileAttachment_validateParameters(parameters) {
  validateStringParam("uri", parameters.uri);
  throwOnOutOfRange(parameters.uri.length, 0, MaxUrlLength, "uri");
  validateAddFileAttachmentApis(parameters.name);
}
// CONCATENATED MODULE: ./src/methods/addBase64FileAttachment.ts






function addBase64FileAttachment(base64String, name) {
  var args = [];

  for (var _i = 2; _i < arguments.length; _i++) {
    args[_i - 2] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "item.addBase64FileAttachmentAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var isInline = false;

  if (!!commonParameters.options) {
    isInline = !!commonParameters.options.isInline;
  }

  var parameters = {
    base64String: base64String,
    name: name,
    isInline: isInline,
    __timeout__: AddItemAttachmentClientEndPointTimeoutInMilliseconds
  };
  addBase64FileAttachment_validateParameters(parameters);
  standardInvokeHostMethod(148, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}

function addBase64FileAttachment_validateParameters(parameters) {
  validateStringParam("base64Encoded", parameters.base64String);
  validateAddFileAttachmentApis(parameters.name);
}
// CONCATENATED MODULE: ./src/methods/addItemAttachment.ts








function addItemAttachment(itemId, name) {
  var args = [];

  for (var _i = 2; _i < arguments.length; _i++) {
    args[_i - 2] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "item.addItemAttachmentAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    itemId: itemId,
    name: name
  };
  addItemAttachment_validateParameters(parameters);
  standardInvokeHostMethod(19, commonParameters.asyncContext, commonParameters.callback, {
    itemId: getItemIdBasedOnHost(parameters.itemId),
    name: parameters.name,
    __timeout__: AddItemAttachmentClientEndPointTimeoutInMilliseconds
  }, undefined);
}

function addItemAttachment_validateParameters(parameters) {
  validateStringParam("itemId", parameters.itemId);
  validateStringParam("attachmentName", parameters.name);
  throwOnOutOfRange(parameters.itemId.length, 0, MaxItemIdLength, "itemId");
  throwOnOutOfRange(parameters.name.length, 0, MaxAttachmentNameLength, "attachmentName");
}
// CONCATENATED MODULE: ./src/methods/close.ts

function close_close() {
  standardInvokeHostMethod(41, undefined, undefined, undefined, undefined);
}
// CONCATENATED MODULE: ./src/methods/getAttachments.ts




function getAttachments_getAttachments() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "item.getAttachmentsAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  standardInvokeHostMethod(149, commonParameters.asyncContext, commonParameters.callback, undefined, CustomJsonAttachmentsResponse);
}
// CONCATENATED MODULE: ./src/methods/getSelectedData.ts





function getSelectedData(coercionType) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "item.getSelectedDataAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  var parameters = {
    coercionType: getCoercionTypeFromString(coercionType)
  };

  if (parameters.coercionType === undefined) {
    throw createArgumentError("coercionType");
  }

  standardInvokeHostMethod(28, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}
// CONCATENATED MODULE: ./src/methods/removeAttachment.ts






function removeAttachment(attachmentId) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "item.removeAttachmentAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    attachmentIndex: attachmentId
  };
  removeAttachment_validateParameters(parameters);
  standardInvokeHostMethod(20, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}

function removeAttachment_validateParameters(parameters) {
  validateStringParam("attachmentId", parameters.attachmentIndex);
  throwOnOutOfRange(parameters.attachmentIndex.length, 0, MaxRemoveIdLength, "attachmentId");
}
// CONCATENATED MODULE: ./src/methods/save.ts



function save() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "item.saveAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  standardInvokeHostMethod(32, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/validation/validateRecipientParameters.ts




function validateRecipientParameters(parameters) {
  if (Array.isArray(parameters.recipientArray)) {
    if (parameters.recipientArray.length > recipientsLimit) {
      throw createArgumentOutOfRange("recipients", parameters.recipientArray.length);
    }

    var validatedRecipients = parameters.recipientArray.map(function (recipient) {
      if (isNullOrUndefined(recipient)) {
        throw createArgumentError("recipients");
      }

      if (typeof recipient === "string") {
        throwOnInvalidDisplayNameOrEmail(recipient, recipient);
        return createEmailAddressForHost(recipient, recipient);
      } else if (typeof recipient === "object") {
        throwOnInvalidDisplayNameOrEmail(recipient.displayName, recipient.emailAddress);
        return createEmailAddressForHost(recipient.displayName, recipient.emailAddress);
      } else {
        throw createArgumentError("recipients");
      }
    });
    parameters.recipientArray = validatedRecipients;
  } else {
    throw createArgumentError("recipients");
  }
}

function throwOnInvalidDisplayNameOrEmail(displayName, email) {
  if (!displayName && !email) {
    throw createArgumentError("recipients");
  } else if (typeof displayName === "string" && displayName.length > displayNameLengthLimit) {
    throw createArgumentOutOfRange("recipients", displayName.length, getString("l_DisplayNameTooLong_Text"));
  } else if (typeof email === "string" && email.length > maxSmtpLength) {
    throw createArgumentOutOfRange("recipients", email.length, getString("l_EmailAddressTooLong_Text"));
  } else if (typeof displayName !== "string" && typeof email !== "string") {
    throw createArgumentError("recipients");
  }
}

function createEmailAddressForHost(displayName, email) {
  return {
    address: email,
    name: displayName
  };
}
// CONCATENATED MODULE: ./src/methods/addRecipients.ts





function addRecipients(namespace) {
  return function (recipientArray) {
    var args = [];

    for (var _i = 1; _i < arguments.length; _i++) {
      args[_i - 1] = arguments[_i];
    }

    checkPermissionsAndThrow(2, namespace + ".addAsync");
    var commonParameters = parseCommonArgs(args, false, false);
    var parameters = {
      recipientField: RecipientFields[namespace],
      recipientArray: recipientArray
    };
    validateRecipientParameters(parameters);
    standardInvokeHostMethod(22, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
  };
}
// CONCATENATED MODULE: ./src/methods/getRecipients.ts





function getRecipients(namespace) {
  return function () {
    var args = [];

    for (var _i = 0; _i < arguments.length; _i++) {
      args[_i] = arguments[_i];
    }

    checkPermissionsAndThrow(1, namespace + ".getAsync");
    var commonParameters = parseCommonArgs(args, true, false);
    standardInvokeHostMethod(15, commonParameters.asyncContext, commonParameters.callback, {
      recipientField: RecipientFields[namespace]
    }, getRecipients_format);
  };
}

function getRecipients_format(rawInput) {
  if (rawInput === null || rawInput === undefined) {
    return [];
  }

  return rawInput.map(function (input) {
    return createEmailAddressDetails(input);
  });
}
// CONCATENATED MODULE: ./src/methods/setRecipients.ts





function setRecipients(namespace) {
  return function (recipientArray) {
    var args = [];

    for (var _i = 1; _i < arguments.length; _i++) {
      args[_i - 1] = arguments[_i];
    }

    checkPermissionsAndThrow(2, namespace + ".setAsync");
    var commonParameters = parseCommonArgs(args, false, false);
    var parameters = {
      recipientField: RecipientFields[namespace],
      recipientArray: recipientArray
    };
    validateRecipientParameters(parameters);
    standardInvokeHostMethod(21, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
  };
}
// CONCATENATED MODULE: ./src/api/getRecipientsSurface.ts




function getRecipientsSurface(namespace) {
  return objectDefine({}, {
    addAsync: addRecipients(namespace),
    getAsync: getRecipients(namespace),
    setAsync: setRecipients(namespace)
  });
}
// CONCATENATED MODULE: ./src/methods/getFrom.ts





function getFrom(namespace) {
  return function () {
    var args = [];

    for (var _i = 0; _i < arguments.length; _i++) {
      args[_i] = arguments[_i];
    }

    checkPermissionsAndThrow(1, namespace + ".getAsync");
    var commonParameters = parseCommonArgs(args, true, false);
    standardInvokeHostMethod(107, commonParameters.asyncContext, commonParameters.callback, undefined, getFrom_format);
  };
}

function getFrom_format(rawInput) {
  return isNullOrUndefined(rawInput) ? null : createEmailAddressDetails(rawInput);
}
// CONCATENATED MODULE: ./src/api/getFromSurface.ts


function getFromSurface(namespace) {
  return objectDefine({}, {
    getAsync: getFrom(namespace)
  });
}
// CONCATENATED MODULE: ./src/validation/validateInternetHeaders.ts



function validateInternetHeaderArray(internetHeaderArray) {
  if (isNullOrUndefined(internetHeaderArray)) {
    throw createArgumentError("internetHeaders");
  }

  if (!Array.isArray(internetHeaderArray)) {
    throw createArgumentTypeError("internetHeaders", typeof internetHeaderArray, "Array");
  }

  if (internetHeaderArray.length === 0) {
    throw createArgumentError("internetHeaders");
  }

  for (var _i = 0, internetHeaderArray_1 = internetHeaderArray; _i < internetHeaderArray_1.length; _i++) {
    var internetHeader = internetHeaderArray_1[_i];
    validateStringParam("internetHeaders", internetHeader);
  }
}
// CONCATENATED MODULE: ./src/methods/removeInternetHeaders.ts




function removeInternetHeaders(internetHeaderKeys) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "internetHeaders.removeAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    internetHeaderKeys: internetHeaderKeys
  };
  removeInternetHeaders_validateParameters(parameters);
  standardInvokeHostMethod(153, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}

function removeInternetHeaders_validateParameters(parameters) {
  validateInternetHeaderArray(parameters.internetHeaderKeys);
}
// CONCATENATED MODULE: ./src/methods/getInternetHeaders.ts




function getInternetHeaders(internetHeaderKeys) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "internetHeaders.getAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  var parameters = {
    internetHeaderKeys: internetHeaderKeys
  };
  getInternetHeaders_validateParameters(parameters);
  standardInvokeHostMethod(151, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}

function getInternetHeaders_validateParameters(parameters) {
  validateInternetHeaderArray(parameters.internetHeaderKeys);
}
// CONCATENATED MODULE: ./src/methods/setInternetHeaders.ts







var InternetHeadersLimit = 998;
function setInternetHeaders(internetHeaderNameValuePairs) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "internetHeaders.setAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    internetHeaderNameValuePairs: internetHeaderNameValuePairs
  };
  setInternetHeaders_validateParameters(parameters);
  standardInvokeHostMethod(152, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}

function setInternetHeaders_validateParameters(parameters) {
  if (isNullOrUndefined(parameters.internetHeaderNameValuePairs)) {
    throw createNullArgumentError("internetHeaders");
  }

  var keys = Object.keys(parameters.internetHeaderNameValuePairs);

  if (keys.length === 0) {
    throw createArgumentError("internetHeaders");
  }

  for (var _i = 0, keys_1 = keys; _i < keys_1.length; _i++) {
    var key = keys_1[_i];
    var value = parameters.internetHeaderNameValuePairs[key];
    validateStringParam("internetHeaders", key);

    if (!(typeof value === "string")) {
      throw createArgumentTypeError("internetHeaders", typeof value, "string");
    }

    throwOnOutOfRange(key.length + value.length, 0, InternetHeadersLimit, key);
  }
}
// CONCATENATED MODULE: ./src/api/getInternetHeadersSurface.ts




function getInternetHeadersSurface(isCompose) {
  var internetHeaders = objectDefine({}, {
    getAsync: getInternetHeaders
  });

  if (isCompose) {
    objectDefine(internetHeaders, {
      removeAsync: removeInternetHeaders,
      setAsync: setInternetHeaders
    });
  }

  return internetHeaders;
}
// CONCATENATED MODULE: ./src/methods/getSubject.ts



function getSubject() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "subject.getAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  standardInvokeHostMethod(18, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/methods/setSubject.ts





var MaximumSubjectLength = 255;
function setSubject(subject) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "subject.setAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    subject: subject
  };
  setSubject_validateParameters(parameters);
  standardInvokeHostMethod(17, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}

function setSubject_validateParameters(parameters) {
  if (!(typeof parameters.subject === "string")) {
    throw createArgumentTypeError("subject", typeof parameters.subject, "string");
  }

  throwOnOutOfRange(parameters.subject.length, 0, MaximumSubjectLength, "subject");
}
// CONCATENATED MODULE: ./src/api/getSubjectSurface.ts



function getSubjectSurface() {
  return objectDefine({}, {
    getAsync: getSubject,
    setAsync: setSubject
  });
}
// CONCATENATED MODULE: ./src/methods/getItemId.ts



function getItemId() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "item.getItemIdAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  standardInvokeHostMethod(164, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/methods/getComposeType.ts




function getComposeType() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "item.getComposeTypeAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  // Kwansong edit
  // checkFeatureEnabledAndThrow(Features.signature, "getComposeTypeAsync");
  standardInvokeHostMethod(174, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/methods/isClientSignatureEnabled.ts




function isClientSignatureEnabled() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "isClientSignatureEnabledAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  checkFeatureEnabledAndThrow(Features.signature, "isClientSignatureEnabledAsync");
  standardInvokeHostMethod(175, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/methods/disableClientSignature.ts




function disableClientSignature() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "disableClientSignatureAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  checkFeatureEnabledAndThrow(Features.signature, "disableClientSignatureAsync");
  standardInvokeHostMethod(176, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/methods/getState.ts




function getState() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "state.getAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  checkFeatureEnabledAndThrow(Features.state, "state.getAsync");
  standardInvokeHostMethod(186, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/methods/setState.ts





function setState(data) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "state.setAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    data: data
  };
  checkFeatureEnabledAndThrow(Features.state, "state.setAsync");
  setState_validateParameters(parameters);
  standardInvokeHostMethod(185, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}

function setState_validateParameters(parameters) {
  validateStringParam("data", parameters.data);
}
// CONCATENATED MODULE: ./src/api/getStateSurface.ts



function getStateSurface() {
  return objectDefine({}, {
    getAsync: getState,
    setAsync: setState
  });
}
// CONCATENATED MODULE: ./src/api/getMessageCompose.ts



























function getMessageCompose() {
  var messageCompose = objectDefine({}, {
    bcc: getRecipientsSurface("bcc"),
    body: getBodySurface(true),
    categories: getCategoriesSurface(),
    cc: getRecipientsSurface("cc"),
    conversationId: getInitialDataProp("conversationId"),
    from: getFromSurface("from"),
    internetHeaders: getInternetHeadersSurface(true),
    itemType: "message",
    notificationMessages: getNotificationMessageSurface(),
    seriesId: getInitialDataProp("seriesId"),
    subject: getSubjectSurface(),
    to: getRecipientsSurface("to"),
    addFileAttachmentAsync: addFileAttachment,
    addFileAttachmentFromBase64Async: addBase64FileAttachment,
    addItemAttachmentAsync: addItemAttachment,
    close: close_close,
    getAttachmentsAsync: getAttachments_getAttachments,
    getAttachmentContentAsync: getAttachmentContent,
    getInitializationContextAsync: getInitializationContext,
    getItemIdAsync: getItemId,
    getSelectedDataAsync: getSelectedData,
    loadCustomPropertiesAsync: loadCustomProperties,
    removeAttachmentAsync: removeAttachment,
    saveAsync: save,
    setSelectedDataAsync: setSelectedData(29),
    delayDeliveryTime: getDelayDeliverySurface(true),
    getComposeTypeAsync: getComposeType,
    isClientSignatureEnabledAsync: isClientSignatureEnabled,
    disableClientSignatureAsync: disableClientSignature,
    state: getStateSurface()
  });
  return messageCompose;
}
// CONCATENATED MODULE: ./src/validation/validateEnhancedLocation.ts




function validateLocationIdentifiers(locationIdentifiers) {
  if (isNullOrUndefined(locationIdentifiers)) {
    throw createNullArgumentError("locationIdentifier");
  }

  if (!Array.isArray(locationIdentifiers)) {
    throw createArgumentTypeError("locationIdentifier", typeof locationIdentifiers, "Array");
  }

  if (locationIdentifiers.length === 0) {
    throw createArgumentError("locationIdentifier");
  }

  for (var _i = 0, locationIdentifiers_1 = locationIdentifiers; _i < locationIdentifiers_1.length; _i++) {
    var locationIdentifier = locationIdentifiers_1[_i];
    validateLocationIdentifier(locationIdentifier);
  }
}

function validateLocationIdentifier(locationIdentifier) {
  if (isNullOrUndefined(locationIdentifier) || isNullOrUndefined(locationIdentifier.id) || isNullOrUndefined(locationIdentifier.type)) {
    throw createNullArgumentError("locationIdentifier");
  }

  if (locationIdentifier.type === MailboxEnums.LocationType.Room || locationIdentifier.type === MailboxEnums.LocationType.Custom) {
    validateIdParameter(locationIdentifier.id, locationIdentifier.type);
  } else {
    throw createArgumentError("type");
  }
}

function validateIdParameter(id, type) {
  if (id === "") {
    throw createArgumentError("id");
  }

  if (type === MailboxEnums.LocationType.Room) {
    if (id.length > maxSmtpLength) {
      throw createArgumentError("id");
    }
  }
}
// CONCATENATED MODULE: ./src/methods/addEnhancedLocations.ts




function addEnhancedLocations(enhancedLocations) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "enhancedLocations.addAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    enhancedLocations: enhancedLocations
  };
  addEnhancedLocations_validateParameters(parameters);
  standardInvokeHostMethod(155, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}

function addEnhancedLocations_validateParameters(parameters) {
  validateLocationIdentifiers(parameters.enhancedLocations);
}
// CONCATENATED MODULE: ./src/methods/getEnhancedLocations.ts



function getEnhancedLocations() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "enhancedLocations.getAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  standardInvokeHostMethod(154, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/methods/removeEnhancedLocations.ts




function removeEnhancedLocations(enhancedLocations) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "enhancedLocations.removeAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    enhancedLocations: enhancedLocations
  };
  removeEnhancedLocations_validateParameters(parameters);
  standardInvokeHostMethod(156, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}

function removeEnhancedLocations_validateParameters(parameters) {
  validateLocationIdentifiers(parameters.enhancedLocations);
}
// CONCATENATED MODULE: ./src/api/getEnhancedLocationSurface.ts




function getEnhancedLocationsSurface(isCompose) {
  var enhancedLocations = objectDefine({}, {
    getAsync: getEnhancedLocations
  });

  if (isCompose) {
    objectDefine(enhancedLocations, {
      addAsync: addEnhancedLocations,
      removeAsync: removeEnhancedLocations
    });
  }

  return enhancedLocations;
}
// CONCATENATED MODULE: ./src/api/getAppointmentRead.ts














function getAppointmentRead() {
  var organizer = getInitialDataProp("organizer");
  var dateTimeCreated = getInitialDataProp("dateTimeCreated");
  var dateTimeModified = getInitialDataProp("dateTimeModified");
  var end = getInitialDataProp("end");
  var start = getInitialDataProp("start");
  var appointmentRead = objectDefine({}, {
    attachments: CustomJsonAttachmentsResponse(getInitialDataProp("attachments")),
    body: getBodySurface(false),
    categories: getCategoriesSurface(),
    dateTimeCreated: dateTimeCreated ? new Date(dateTimeCreated) : undefined,
    dateTimeModified: dateTimeModified ? new Date(dateTimeModified) : undefined,
    end: end ? new Date(end) : undefined,
    enhancedLocation: getEnhancedLocationsSurface(false),
    itemClass: getInitialDataProp("itemClass"),
    itemId: getInitialDataProp("id"),
    itemType: "appointment",
    location: getInitialDataProp("location"),
    normalizedSubject: getInitialDataProp("normalizedSubject"),
    notificationMessages: getNotificationMessageSurface(),
    optionalAttendees: (getInitialDataProp("cc") || []).map(createEmailAddressDetails),
    organizer: organizer ? createEmailAddressDetails(organizer) : undefined,
    recurrence: copyRecurrenceObjectConvertSeriesTimeJson(getInitialDataProp("recurrence")),
    requiredAttendees: (getInitialDataProp("to") || []).map(createEmailAddressDetails),
    start: start ? new Date(start) : undefined,
    seriesId: getInitialDataProp("seriesId"),
    subject: getInitialDataProp("subject"),
    displayReplyForm: displayReplyForm,
    displayReplyFormAsync: displayReplyFormAsync,
    displayReplyAllForm: displayReplyAllForm,
    displayReplyAllFormAsync: displayReplyAllFormAsync,
    getAttachmentContentAsync: getAttachmentContent,
    getEntities: Entities_getEntities,
    getEntitiesByType: Entities_getEntitiesByType,
    getFilteredEntitiesByName: Entities_getFilteredEntitiesByName,
    getInitializationContextAsync: getInitializationContext,
    getRegExMatches: Entities_getRegExMatches,
    getRegExMatchesByName: Entities_getRegExMatchesByName,
    getSelectedEntities: Entities_getSelectedEntities,
    getSelectedRegExMatches: Entities_getSelectedRegExMatches,
    loadCustomPropertiesAsync: loadCustomProperties,
    isAllDayEvent: getInitialDataProp("isAllDayEvent"),
    sensitivity: getInitialDataProp("sensitivity")
  });
  return appointmentRead;
}
// CONCATENATED MODULE: ./src/validation/timeConstants.ts
var TimeType;

(function (TimeType) {
  TimeType[TimeType["start"] = 1] = "start";
  TimeType[TimeType["end"] = 2] = "end";
})(TimeType || (TimeType = {}));
// CONCATENATED MODULE: ./src/methods/getTime.ts




function getTime(namespace) {
  return function () {
    var args = [];

    for (var _i = 0; _i < arguments.length; _i++) {
      args[_i] = arguments[_i];
    }

    checkPermissionsAndThrow(1, namespace + ".getAsync");
    var commonParameters = parseCommonArgs(args, true, false);
    standardInvokeHostMethod(24, commonParameters.asyncContext, commonParameters.callback, {
      TimeProperty: TimeType[namespace]
    }, getTime_format);
  };
}

function getTime_format(rawInput) {
  var ticks = rawInput;
  return new Date(ticks);
}
// CONCATENATED MODULE: ./src/methods/setTime.ts





var maxTime = 8640000000000000;
var minTime = -8640000000000000;
function setTime(namespace) {
  return function (date) {
    var args = [];

    for (var _i = 1; _i < arguments.length; _i++) {
      args[_i - 1] = arguments[_i];
    }

    checkPermissionsAndThrow(2, namespace + ".setAsync");
    var commonParameters = parseCommonArgs(args, false, false);
    var parameters = {
      date: date
    };
    setTime_validateParameters(parameters);
    standardInvokeHostMethod(25, commonParameters.asyncContext, commonParameters.callback, {
      TimeProperty: TimeType[namespace],
      time: parameters.date.getTime()
    }, undefined);
  };
}

function setTime_validateParameters(parameters) {
  if (!(parameters.date instanceof Date)) {
    throw createArgumentTypeError("dateTime", typeof parameters.date, typeof Date);
  }

  if (isNaN(parameters.date.getTime())) {
    throw createArgumentError("dateTime");
  }

  if (parameters.date.getTime() < minTime || parameters.date.getTime() > maxTime) {
    throw createArgumentOutOfRange("dateTime");
  }
}
// CONCATENATED MODULE: ./src/api/getTimeSurface.ts



function getTimeSurface(namespace) {
  return objectDefine({}, {
    getAsync: getTime(namespace),
    setAsync: setTime(namespace)
  });
}
// CONCATENATED MODULE: ./src/methods/getLocation.ts



function getLocation() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "location.getAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  standardInvokeHostMethod(26, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/methods/setLocation.ts






var MaximumLocationLength = 255;
function setLocation(location) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "location.setAsync");
  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    location: location
  };
  setLocation_validateParameters(parameters);
  standardInvokeHostMethod(27, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}

function setLocation_validateParameters(parameters) {
  if (!isNullOrUndefined(parameters.location)) {
    if (!(typeof parameters.location === "string")) {
      throw createArgumentTypeError("location", typeof parameters.location, "string");
    }

    throwOnOutOfRange(parameters.location.length, 0, MaximumLocationLength, "location");
  } else {
    throw createNullArgumentError("location");
  }
}
// CONCATENATED MODULE: ./src/api/getLocationSurface.ts



function getLocationSurface() {
  return objectDefine({}, {
    getAsync: getLocation,
    setAsync: setLocation
  });
}
// CONCATENATED MODULE: ./src/methods/getRecurrence.ts




function getRecurrence() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "recurrenceProperties.getAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  standardInvokeHostMethod(103, commonParameters.asyncContext, commonParameters.callback, undefined, seriesTimeJsonConverter);
}
function seriesTimeJsonConverter(rawInput) {
  if (rawInput !== null) {
    if (rawInput.seriesTimeJson !== null) {
      var seriesTime = new SeriesTime_SeriesTime();
      seriesTime.importFromSeriesTimeJsonObject(rawInput.seriesTimeJson);
      delete rawInput.seriesTimeJson;
      rawInput.seriesTime = seriesTime;
    }
  }

  return rawInput;
}
// CONCATENATED MODULE: ./src/validation/validateRecurrenceObject.ts






function validateRecurrenceObject(recurrenceObject) {
  if (isNullOrUndefined(recurrenceObject)) {
    return;
  }

  recurrenceObject = recurrenceObject;

  if (isNullOrUndefined(recurrenceObject.recurrenceType)) {
    throw createNullArgumentError(RecurrenceTypeKey);
  }

  if (isNullOrUndefined(recurrenceObject.seriesTime)) {
    throw createNullArgumentError(SeriesTimeKey);
  }

  if (!(recurrenceObject.seriesTime instanceof SeriesTime_SeriesTime) || !recurrenceObject.seriesTime.isValid()) {
    throw createArgumentError(SeriesTimeKey);
  }

  if (!recurrenceObject.seriesTime.isEndAfterStart()) {
    throw createArgumentError(SeriesTimeKey, getString("l_InvalidEventDates_Text"));
  }

  throwOnInvalidRecurrenceType(recurrenceObject.recurrenceType);

  if (recurrenceObject.recurrenceType !== MailboxEnums.RecurrenceType.Weekday) {
    if (isNullOrUndefined(recurrenceObject.recurrenceProperties)) {
      throw createNullArgumentError(RecurrenceTypeKey);
    }
  }

  if (!isNullOrUndefined(recurrenceObject.recurrenceTimeZone)) {
    if (isNullOrUndefined(recurrenceObject.recurrenceTimeZone.name)) {
      throw createNullArgumentError(RecurrenceTimeZoneName);
    }

    if (typeof recurrenceObject.recurrenceTimeZone.name !== "string") {
      throw createArgumentTypeError(RecurrenceTimeZoneName, typeof recurrenceObject.recurrenceTimeZone.name, "string");
    }
  }

  if (recurrenceObject.recurrenceType === MailboxEnums.RecurrenceType.Daily) {
    throwOnInvalidDailyRecurrence(recurrenceObject.recurrenceProperties);
  } else if (recurrenceObject.recurrenceType === MailboxEnums.RecurrenceType.Weekly) {
    throwOnInvalidWeeklyRecurrence(recurrenceObject.recurrenceProperties);
  } else if (recurrenceObject.recurrenceType === MailboxEnums.RecurrenceType.Monthly) {
    throwOnInvalidMonthlyRecurrence(recurrenceObject.recurrenceProperties);
  } else if (recurrenceObject.recurrenceType === MailboxEnums.RecurrenceType.Yearly) {
    throwOnInvalidYearlyRecurrence(recurrenceObject.recurrenceProperties);
  }
}

function throwOnInvalidRecurrenceType(recurrenceType) {
  if (recurrenceType !== MailboxEnums.RecurrenceType.Daily && recurrenceType !== MailboxEnums.RecurrenceType.Weekly && recurrenceType !== MailboxEnums.RecurrenceType.Weekday && recurrenceType !== MailboxEnums.RecurrenceType.Yearly && recurrenceType !== MailboxEnums.RecurrenceType.Monthly) {
    throw createArgumentError(RecurrenceTypeKey);
  }
}

function throwOnInvalidRecurrenceInterval(recurrenceProperties) {
  if (isNullOrUndefined(recurrenceProperties.interval)) {
    throw createNullArgumentError(IntervalKey);
  }

  if (typeof recurrenceProperties.interval !== "number") {
    throw createArgumentTypeError(IntervalKey, typeof recurrenceProperties.interval, "number");
  }

  if (recurrenceProperties.interval <= 0) {
    throw createArgumentError(IntervalKey);
  }
}

function throwOnInvalidDailyRecurrence(recurrenceProperties) {
  throwOnInvalidRecurrenceInterval(recurrenceProperties);
}

function throwOnInvalidWeeklyRecurrence(recurrenceProperties) {
  throwOnInvalidRecurrenceInterval(recurrenceProperties);

  if (isNullOrUndefined(recurrenceProperties.days)) {
    throw createArgumentTypeError(DaysKey);
  }

  if (!Array.isArray(recurrenceProperties.days)) {
    throw createArgumentTypeError(DaysKey);
  }

  throwOnInvalidDaysArray(recurrenceProperties.days);

  if (!isNullOrUndefined(recurrenceProperties.firstDayOfWeek)) {
    if (typeof recurrenceProperties.firstDayOfWeek !== "string") {
      throw createArgumentTypeError(FirstDayOfWeekKey);
    }

    if (!verifyDays(recurrenceProperties.firstDayOfWeek, false)) {
      throw createArgumentError(FirstDayOfWeekKey);
    }
  }
}

function throwOnInvalidDaysArray(daysArray) {
  for (var i = 0; i < daysArray.length; i++) {
    if (!verifyDays(daysArray[i], false)) {
      throw createArgumentError(DaysKey);
    }
  }
}

function verifyDays(dayEnum, checkGroupedDays) {
  var fRegularDay = dayEnum === MailboxEnums.Days.Mon || dayEnum === MailboxEnums.Days.Tue || dayEnum === MailboxEnums.Days.Wed || dayEnum === MailboxEnums.Days.Thu || dayEnum === MailboxEnums.Days.Fri || dayEnum === MailboxEnums.Days.Sat || dayEnum === MailboxEnums.Days.Sun;

  if (checkGroupedDays) {
    var fGroupedDay = dayEnum === MailboxEnums.Days.WeekendDay || dayEnum === MailboxEnums.Days.Weekday || dayEnum === MailboxEnums.Days.Day;
    return fGroupedDay || fRegularDay;
  } else {
    return fRegularDay;
  }
}

function throwOnInvalidMonthlyRecurrence(recurrenceProperties) {
  if (isNullOrUndefined(recurrenceProperties.interval)) {
    throw createNullArgumentError(IntervalKey);
  }

  if (typeof recurrenceProperties.interval !== "number") {
    throw createArgumentTypeError(IntervalKey, typeof recurrenceProperties.interval, "number");
  }

  if (!isNullOrUndefined(recurrenceProperties.dayOfMonth)) {
    if (typeof recurrenceProperties.dayOfMonth !== "number") {
      throw createArgumentTypeError(DayOfMonthKey, typeof recurrenceProperties.dayOfMonth, "number");
    }

    throwOnInvalidDayOfMonth(recurrenceProperties.dayOfMonth);
  } else if (!isNullOrUndefined(recurrenceProperties.dayOfWeek) && !isNullOrUndefined(recurrenceProperties.weekNumber)) {
    if (typeof recurrenceProperties.dayOfWeek !== "string") {
      throw createArgumentTypeError(DayOfWeekKey, typeof recurrenceProperties.dayOfWeek, "string");
    }

    if (!verifyDays(recurrenceProperties.dayOfWeek, true)) {
      throw createArgumentError(DayOfWeekKey);
    }

    if (typeof recurrenceProperties.weekNumber !== "string") {
      throw createArgumentTypeError(WeekNumberKey, typeof recurrenceProperties.weekNumber, "string");
    }

    throwOnInvalidWeekNumber(recurrenceProperties.weekNumber);
  } else {
    throw createArgumentError(undefined, getString("l_Recurrence_Error_Properties_Invalid_Text"));
  }
}

function throwOnInvalidWeekNumber(weekNumber) {
  if (weekNumber !== MailboxEnums.WeekNumber.First && weekNumber !== MailboxEnums.WeekNumber.Second && weekNumber !== MailboxEnums.WeekNumber.Third && weekNumber !== MailboxEnums.WeekNumber.Fourth && weekNumber !== MailboxEnums.WeekNumber.Last) {
    throw createArgumentError(WeekNumberKey);
  }
}

function throwOnInvalidDayOfMonth(iDayOfMonth) {
  if (iDayOfMonth < 1 || iDayOfMonth > 31) {
    throw createArgumentError(DayOfMonthKey);
  }
}

function throwOnInvalidYearlyRecurrence(recurrenceProperties) {
  if (isNullOrUndefined(recurrenceProperties.interval)) {
    throw createNullArgumentError(IntervalKey);
  }

  if (typeof recurrenceProperties.interval !== "number") {
    throw createArgumentTypeError(IntervalKey, typeof recurrenceProperties.interval, "number");
  }

  if (isNullOrUndefined(recurrenceProperties.month)) {
    throw createNullArgumentError(MonthKey);
  }

  if (typeof recurrenceProperties.month !== "string") {
    throw createArgumentTypeError(MonthKey, typeof recurrenceProperties.month, "string");
  }

  throwOnInvalidMonth(recurrenceProperties.month);

  if (!isNullOrUndefined(recurrenceProperties.dayOfMonth)) {
    if (typeof recurrenceProperties.dayOfMonth !== "number") {
      throw createArgumentTypeError(DayOfMonthKey, typeof recurrenceProperties.dayOfMonth, "number");
    }

    throwOnInvalidDayOfMonth(recurrenceProperties.dayOfMonth);
  } else if (!isNullOrUndefined(recurrenceProperties.weekNumber) && !isNullOrUndefined(recurrenceProperties.dayOfWeek)) {
    if (typeof recurrenceProperties.dayOfWeek !== "string") {
      throw createArgumentTypeError(DayOfWeekKey, typeof recurrenceProperties.dayOfWeek, "string");
    }

    if (!verifyDays(recurrenceProperties.dayOfWeek, true)) {
      throw createArgumentError(DayOfWeekKey);
    }

    if (typeof recurrenceProperties.weekNumber !== "string") {
      throw createArgumentTypeError(WeekNumberKey, typeof recurrenceProperties.weekNumber, "string");
    }

    throwOnInvalidWeekNumber(recurrenceProperties.weekNumber);
  } else {
    throw createArgumentError(undefined, getString("l_Recurrence_Error_Properties_Invalid_Text"));
  }
}

function throwOnInvalidMonth(month) {
  if (month !== MailboxEnums.Month.Jan && month !== MailboxEnums.Month.Feb && month !== MailboxEnums.Month.Mar && month !== MailboxEnums.Month.Apr && month !== MailboxEnums.Month.May && month !== MailboxEnums.Month.Jun && month !== MailboxEnums.Month.Jul && month !== MailboxEnums.Month.Aug && month !== MailboxEnums.Month.Sep && month !== MailboxEnums.Month.Oct && month !== MailboxEnums.Month.Nov && month !== MailboxEnums.Month.Dec) {
    throw createArgumentError(MonthKey);
  }
}
// CONCATENATED MODULE: ./src/methods/setRecurrence.ts









function setRecurrence(recurrencePattern) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "recurrenceProperties.setAsync");
  var seriesId = getAppointmentCompose().seriesId;

  if (!isNullOrUndefined(seriesId) && seriesId.length > 0) {
    throw createArgumentError(undefined, getString("l_Recurrence_Error_Instance_SetAsync_Text"));
  }

  validateRecurrenceObject(recurrencePattern);
  var commonParameters = parseCommonArgs(args, false, false);
  var recurrenceData = convertSeriesTime(recurrencePattern);
  var parameters = {
    recurrenceData: recurrenceData
  };
  standardInvokeHostMethod(104, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}

function convertSeriesTime(recurrencePattern) {
  if (recurrencePattern !== null && recurrencePattern.seriesTime !== null) {
    if (recurrencePattern.seriesTime instanceof SeriesTime_SeriesTime) {
      var recurrencePatternCopy = {
        recurrenceProperties: recurrencePattern.recurrenceProperties,
        recurrenceTimeZone: recurrencePattern.recurrenceTimeZone,
        recurrenceType: recurrencePattern.recurrenceType,
        seriesTimeJson: recurrencePattern.seriesTime.exportToSeriesTimeJson()
      };
      return recurrencePatternCopy;
    }
  }

  return recurrencePattern;
}
// CONCATENATED MODULE: ./src/api/getRecurrenceSurface.ts



function getRecurrenceSurface(isCompose) {
  var recurrence = objectDefine({}, {
    getAsync: getRecurrence
  });

  if (isCompose) {
    objectDefine(recurrence, {
      setAsync: setRecurrence
    });
  }

  return recurrence;
}
// CONCATENATED MODULE: ./src/methods/getAllDayEvent.ts




function getAllDayEvent() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "isAllDayEvent.getAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  checkFeatureEnabledAndThrow(Features.calendarItems, "isAllDayEvent.getAsync");
  standardInvokeHostMethod(169, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/methods/setAllDayEvent.ts






function setAllDayEvent(isAllDayEvent) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "isAllDayEvent.setAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  var parameters = {
    isAllDayEvent: isAllDayEvent
  };
  checkFeatureEnabledAndThrow(Features.calendarItems, "isAllDayEvent.setAsync");
  setAllDayEvent_validateParameters(parameters);
  standardInvokeHostMethod(170, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}

function setAllDayEvent_validateParameters(parameters) {
  if (isNullOrUndefined(parameters.isAllDayEvent)) {
    throw createNullArgumentError("isAllDayEvent");
  }

  if (typeof parameters.isAllDayEvent !== "boolean") {
    throw createArgumentTypeError("isAllDayEvent", typeof parameters.isAllDayEvent, "boolean");
  }
}
// CONCATENATED MODULE: ./src/api/getAllDayEventSurface.ts



function getAllDayEventSurface() {
  return objectDefine({}, {
    getAsync: getAllDayEvent,
    setAsync: setAllDayEvent
  });
}
// CONCATENATED MODULE: ./src/methods/setSensitivity.ts







function setSensitivity(sensitivity) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  checkPermissionsAndThrow(2, "sensitivity.setAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  var parameters = {
    sensitivity: sensitivity
  };
  checkFeatureEnabledAndThrow(Features.calendarItems, "sensitivity.setAsync");
  setSensitivity_validateParameters(parameters);
  standardInvokeHostMethod(172, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}

function setSensitivity_validateParameters(parameters) {
  validateStringParam("sensitivity", parameters.sensitivity);
  throwOnInvalidSensitivityType(parameters.sensitivity);
}

function throwOnInvalidSensitivityType(sensitivity) {
  if (sensitivity !== MailboxEnums.AppointmentSensitivityType.Normal && sensitivity !== MailboxEnums.AppointmentSensitivityType.Personal && sensitivity !== MailboxEnums.AppointmentSensitivityType.Private && sensitivity !== MailboxEnums.AppointmentSensitivityType.Confidential) {
    throw createArgumentError("sensitivity");
  }
}
// CONCATENATED MODULE: ./src/methods/getSensitivity.ts




function getSensitivity() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "sensitivity.getAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  checkFeatureEnabledAndThrow(Features.calendarItems, "sensitivity.getAsync");
  standardInvokeHostMethod(171, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/api/getSensitivitySurface.ts



function getSensitivitySurface() {
  return objectDefine({}, {
    getAsync: getSensitivity,
    setAsync: setSensitivity
  });
}
// CONCATENATED MODULE: ./src/api/getAppointmentCompose.ts






























function getAppointmentCompose() {
  var appointmentCompose = objectDefine({}, {
    body: getBodySurface(true),
    categories: getCategoriesSurface(),
    end: getTimeSurface("end"),
    enhancedLocation: getEnhancedLocationsSurface(true),
    itemType: "appointment",
    location: getLocationSurface(),
    notificationMessages: getNotificationMessageSurface(),
    optionalAttendees: getRecipientsSurface("optionalAttendees"),
    organizer: getFromSurface("organizer"),
    recurrence: getRecurrenceSurface(true),
    requiredAttendees: getRecipientsSurface("requiredAttendees"),
    seriesId: getInitialDataProp("seriesId"),
    start: getTimeSurface("start"),
    subject: getSubjectSurface(),
    addFileAttachmentAsync: addFileAttachment,
    addFileAttachmentFromBase64Async: addBase64FileAttachment,
    addItemAttachmentAsync: addItemAttachment,
    close: close_close,
    getAttachmentsAsync: getAttachments_getAttachments,
    getAttachmentContentAsync: getAttachmentContent,
    getInitializationContextAsync: getInitializationContext,
    getItemIdAsync: getItemId,
    getSelectedDataAsync: getSelectedData,
    loadCustomPropertiesAsync: loadCustomProperties,
    removeAttachmentAsync: removeAttachment,
    saveAsync: save,
    setSelectedDataAsync: setSelectedData(29),
    isAllDayEvent: getAllDayEventSurface(),
    sensitivity: getSensitivitySurface(),
    isClientSignatureEnabledAsync: isClientSignatureEnabled,
    disableClientSignatureAsync: disableClientSignature,
    state: getStateSurface()
  });
  return appointmentCompose;
}
// CONCATENATED MODULE: ./src/methods/registerConsent.ts



var ConsentStateType;

(function (ConsentStateType) {
  ConsentStateType[ConsentStateType["NotResponded"] = 0] = "NotResponded";
  ConsentStateType[ConsentStateType["NotConsented"] = 1] = "NotConsented";
  ConsentStateType[ConsentStateType["Consented"] = 2] = "Consented";
})(ConsentStateType || (ConsentStateType = {}));

function registerConsent(consentState) {
  var parameters = {
    consentState: consentState,
    extensionId: getInitialDataProp("extensionId")
  };
  registerConsent_validateParameters(consentState);
  standardInvokeHostMethod(40, undefined, undefined, parameters, undefined);
}

function registerConsent_validateParameters(consentState) {
  if (consentState !== ConsentStateType.Consented && consentState !== ConsentStateType.NotConsented && consentState !== ConsentStateType.NotResponded) {
    throw createArgumentOutOfRange("consentState");
  }
}
// CONCATENATED MODULE: ./src/methods/navigateToModule.ts





function navigateToModule(moduleName) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    module: moduleName
  };
  navigateToModule_validateParameters(moduleName);

  if (moduleName === MailboxEnums.ModuleType.Addins) {
    if (!!commonParameters.options && !!commonParameters.options.queryString) {
      parameters.queryString = commonParameters.options.queryString;
    } else {
      parameters.queryString = "";
    }
  }

  standardInvokeHostMethod(45, commonParameters.asyncContext, commonParameters.callback, parameters, undefined);
}

function navigateToModule_validateParameters(moduleName) {
  if (isNullOrUndefined(moduleName)) {
    throw createNullArgumentError("module");
  }

  if (moduleName === "") {
    throw createArgumentError("module");
  }

  if (moduleName !== MailboxEnums.ModuleType.Addins) {
    throw createArgumentError("module");
  }
}
// CONCATENATED MODULE: ./src/methods/recordDataPoint.ts



function recordDataPoint(data) {
  if (isNullOrUndefined(data)) {
    throw createNullArgumentError("data");
  }

  standardInvokeHostMethod(402, undefined, undefined, data, undefined);
}
// CONCATENATED MODULE: ./src/methods/recordTrace.ts



function recordTrace(data) {
  if (isNullOrUndefined(data)) {
    throw createNullArgumentError("data");
  }

  standardInvokeHostMethod(401, undefined, undefined, data, undefined);
}
// CONCATENATED MODULE: ./src/methods/trackCtq.ts



function trackCtq(data) {
  if (isNullOrUndefined(data)) {
    throw createNullArgumentError("data");
  }

  standardInvokeHostMethod(400, undefined, undefined, data, undefined);
}
// CONCATENATED MODULE: ./src/methods/logTelemetry.ts



function logTelemetry(data) {
  if (isNullOrUndefined(data)) {
    throw createNullArgumentError("data");
  }

  standardInvokeHostMethod(163, undefined, undefined, {
    telemetryData: data
  }, undefined);
}
// CONCATENATED MODULE: ./src/utils/convertToLocalClientTime.ts
var __assign = undefined && undefined.__assign || function () {
  __assign = Object.assign || function (t) {
    for (var s, i = 1, n = arguments.length; i < n; i++) {
      s = arguments[i];

      for (var p in s) {
        if (Object.prototype.hasOwnProperty.call(s, p)) t[p] = s[p];
      }
    }

    return t;
  };

  return __assign.apply(this, arguments);
};






function convertToLocalClientTime(timeValue) {
  if (!(timeValue instanceof Date)) {
    throw createArgumentError("timeValue");
  }

  var date = new Date(timeValue.getTime());
  var offset = date.getTimezoneOffset() * -1;

  if (!isNullOrUndefined(getInitialDataProp("timeZoneOffsets"))) {
    date.setUTCMinutes(date.getUTCMinutes() - offset);
    offset = findOffset(date);
    date.setUTCMinutes(date.getUTCMinutes() + offset);
  }

  var retValue = __assign({
    timezoneOffset: offset
  }, dateToDictionary(date));

  return retValue;
}
// CONCATENATED MODULE: ./src/methods/displayPersonaCardAsync.ts




function displayPersonaCardAsync(ewsIdOrEmail) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  var commonParameters = parseCommonArgs(args, false, false);
  var parameters = {
    ewsIdOrEmail: ewsIdOrEmail
  };
  displayPersonaCardAsync_validateParameters(parameters);
  standardInvokeHostMethod(43, commonParameters.asyncContext, commonParameters.callback, {
    ewsIdOrEmail: ewsIdOrEmail.trim()
  }, undefined);
}

function displayPersonaCardAsync_validateParameters(parameters) {
  if (!isNullOrUndefined(parameters.ewsIdOrEmail)) {
    displayPersonaCardAsync_throwOnInvalidItemId(parameters.ewsIdOrEmail);

    if (parameters.ewsIdOrEmail === "") {
      throw createArgumentError("ewsIdOrEmail", "ewsIdOrEmail cannot be empty.");
    }
  } else {
    throw createNullArgumentError("ewsIdOrEmail");
  }
}

function displayPersonaCardAsync_throwOnInvalidItemId(ewsIdOrEmail) {
  if (!(typeof ewsIdOrEmail === "string")) {
    throw createArgumentError("ewsIdOrEmail");
  }
}
// CONCATENATED MODULE: ./src/methods/getSharedProperties.ts



function getSharedProperties() {
  var args = [];

  for (var _i = 0; _i < arguments.length; _i++) {
    args[_i] = arguments[_i];
  }

  checkPermissionsAndThrow(1, "item.getSharedPropertiesAsync");
  var commonParameters = parseCommonArgs(args, true, false);
  standardInvokeHostMethod(108, commonParameters.asyncContext, commonParameters.callback, undefined, undefined);
}
// CONCATENATED MODULE: ./src/utils/addSharedPropertiesSupport.ts





var addSharedPropertiesSupport_addSharedPropertiesSupport = function addSharedPropertiesSupport(target) {
  if (target && getInitialDataProp("isFromSharedFolder") && getHostItemType_getHostItemType() !== HostItemType.ItemLess) {
    objectDefine(target, {
      getSharedPropertiesAsync: getSharedProperties
    });
  }
};
// CONCATENATED MODULE: ./src/api/prepareApiSurface.ts
































var prepareApiSurface_OSF = __webpack_require__(0);

var prepareApiSurface_createMailboxSurface = function createMailboxSurface(target) {
  objectDefine(target, {
    ewsUrl: getInitialDataProp("ewsUrl"),
    restUrl: getInitialDataProp("restUrl"),
    displayAppointmentForm: displayAppointmentForm,
    displayAppointmentFormAsync: displayAppointmentFormAsync,
    displayMessageForm: displayMessageForm,
    displayMessageFormAsync: displayMessageFormAsync,
    displayPersonaCardAsync: displayPersonaCardAsync,
    getCallbackTokenAsync: getCallbackToken,
    getUserIdentityTokenAsync: getUserIdentityToken,
    logTelemetry: logTelemetry,
    makeEwsRequestAsync: makeEwsRequest,
    masterCategories: getMasterCategoriesSurface(),
    navigateToModuleAsync: navigateToModule,
    diagnostics: getDiagnosticsSurface(),
    userProfile: getUserProfileSurface(),
    convertToEwsId: convertToEwsId,
    convertToLocalClientTime: convertToLocalClientTime,
    convertToRestId: convertToRestId,
    convertToUtcClientTime: convertToUtcClientTime,
    RegisterConsentAsync: registerConsent,
    GetIsRead: function GetIsRead() {
      return getInitialDataProp("isRead");
    },
    GetEndPointUrl: function GetEndPointUrl() {
      return getInitialDataProp("endNodeUrl");
    },
    GetConsentMetaData: function GetConsentMetaData() {
      return getInitialDataProp("consentMetadata");
    },
    GetMarketplaceContentMarket: function GetMarketplaceContentMarket() {
      return getInitialDataProp("marketplaceContentMarket");
    },
    GetMarketplaceAssetId: function GetMarketplaceAssetId() {
      return getInitialDataProp("marketplaceAssetId");
    },
    GetExtensionId: function GetExtensionId() {
      return getInitialDataProp("extensionId");
    },
    CloseApp: closeApp,
    recordDataPoint: recordDataPoint,
    recordTrace: recordTrace,
    trackCtq: trackCtq
  });

  if (getHostItemType_getHostItemType() !== HostItemType.MessageCompose && getHostItemType_getHostItemType() !== HostItemType.AppointmentCompose) {
    objectDefine(target, {
      displayNewAppointmentForm: displayNewAppointmentForm,
      displayNewMessageForm: displayNewMessageForm,
      displayNewAppointmentFormAsync: displayNewAppointmentFormAsync,
      displayNewMessageFormAsync: displayNewMessageFormAsync
    });
  }

  return target;
};
var prepareApiSurface_getItem = function getItem() {
  var item = undefined;

  switch (getHostItemType_getHostItemType()) {
    case HostItemType.Message:
      item = getMessageRead();
      break;

    case HostItemType.MessageCompose:
      item = getMessageCompose();
      break;

    case HostItemType.Appointment:
      item = getAppointmentRead();
      break;

    case HostItemType.AppointmentCompose:
      item = getAppointmentCompose();
      break;

    case HostItemType.MeetingRequest:
      item = getMessageRead();
      break;

    default:
      return undefined;
  }

  addSharedPropertiesSupport_addSharedPropertiesSupport(item);
  return item;
};
// CONCATENATED MODULE: ./src/api/OutlookAppOm.ts





var OutlookAppOm_OSF = __webpack_require__(0);

var appInstance;
var whenStringsFinish;
var getInitialDataProp = function getInitialDataProp(key) {
  return appInstance && appInstance.getInitialDataProp(key);
};
var getIsNoItemContextWebExt = function getIsNoItemContextWebExt() {
  return !appInstance || !appInstance.item;
};
var getAppName = function getAppName() {
  return appInstance && appInstance.getAppName();
};

var OutlookAppOm_OutlookAppOm = function () {
  function OutlookAppOm(appContext, targetWindow, appReadyCallback) {
    var _this = this;

    this.displayName = "mailbox";

    this.stringLoadedCallback = function () {
      if (!!_this.appReadyCallback) {
        if (!_this.officeAppContext.get_isDialog()) {
          standardInvokeHostMethod_invokeHostMethod(1, undefined, _this.onInitialDataResponse);
        } else {
          setTimeout(function () {
            return _this.appReadyCallback();
          }, 0);
        }
      }
    };

    this.initialize = function (data) {
      if (data === null || data === undefined) {
        recreateAdditionalGlobalParametersSingleton(true);
        _this.initialData = null;
        _this.item = null;
      } else {
        _this.initialData = data;
        _this.initialData.permissionLevel = calculatePermissionLevel();
        _this.item = prepareApiSurface_getItem();
        var supportsAdditionalParameters = false;
        supportsAdditionalParameters = true;
        recreateAdditionalGlobalParametersSingleton(supportsAdditionalParameters);

        if (typeof data.itemNumber !== "undefined") {
          getAdditionalGlobalParametersSingleton().setCurrentItemNumber(data.itemNumber);
        }
      }
    };

    this.onInitialDataResponse = function (resultCode, data) {
      if (!!resultCode && resultCode !== InvokeResultCode.noError) {
        return;
      }

      _this.initialize(data);

      prepareApiSurface_createMailboxSurface(_this);
      setTimeout(function () {
        return _this.appReadyCallback();
      }, 0);
    };

    this.officeAppContext = appContext;
    this.targetWindow = window;
    this.appReadyCallback = appReadyCallback;
    appInstance = this;

    if (OutlookAppOm.shouldRunInitialDataResponse()) {
      this.onInitialDataResponse(InvokeResultCode.noError, GlobalMailboxInitialData);
    }
  }

  OutlookAppOm.prototype.getAppName = function () {
    var retVal = -1;
    retVal = 8;
    return retVal;
  };

  OutlookAppOm.prototype.getInitialDataProp = function (key) {
    return this.initialData && this.initialData[key];
  };

  OutlookAppOm.prototype.setCurrentItemNumber = function (newItemNumber) {
    getAdditionalGlobalParametersSingleton().setCurrentItemNumber(newItemNumber);
  };

  OutlookAppOm.addAdditionalArgs = function (dispid, hostCallingArgs) {
    return hostCallingArgs;
  };

  OutlookAppOm.shouldRunInitialDataResponse = function () {
    return true;
  };

  return OutlookAppOm;
}();



var calculatePermissionLevel = function calculatePermissionLevel() {
  var HostReadItem = 1;
  var HostReadWriteMailbox = 2;
  var HostReadWriteItem = 3;
  var permissionLevelFromHost = getInitialDataProp("permissionLevel");

  if (permissionLevelFromHost === undefined) {
    return 0;
  }

  switch (permissionLevelFromHost) {
    case HostReadItem:
      return 1;

    case HostReadWriteItem:
      return 2;

    case HostReadWriteMailbox:
      return 3;

    default:
      return 0;
  }
};
// CONCATENATED MODULE: ./src/methods/saveSettingsRequest.ts





var saveSettingsRequest_OSF = __webpack_require__(0);

var settingsMaxNumberOfCharacters = 32 * 1024;
function saveSettingsRequest(data) {
  var args = [];

  for (var _i = 1; _i < arguments.length; _i++) {
    args[_i - 1] = arguments[_i];
  }

  var commonParameters = parseCommonArgs(args, false, false);
  var serializedSettings = saveSettingsRequest_OSF.DDA.SettingsManager.serializeSettings(data);

  if (JSON.stringify(serializedSettings).length > settingsMaxNumberOfCharacters) {
    var asyncResult_1 = createAsyncResult(undefined, saveSettingsRequest_OSF.DDA.AsyncResultEnum.ErrorCode.Failed, 9019, commonParameters.asyncContext, "");

    if (!!commonParameters.callback) {
      setTimeout(function () {
        if (!!commonParameters.callback) commonParameters.callback(asyncResult_1);
      }, 0);
    }

    return;
  }

  if (saveSettingsRequest_OSF.AppName.OutlookWebApp === getAppName()) {
    saveSettingsForOwa(commonParameters, serializedSettings);
  } else {
    saveSettingsForOutlookDesktop(commonParameters, serializedSettings);
  }
}

function saveSettingsForOwa(commonParameters, serializedSettings) {
  standardInvokeHostMethod(404, commonParameters.asyncContext, commonParameters.callback, [serializedSettings], undefined);
}

function saveSettingsForOutlookDesktop(commonParameters, serializedSettings) {
  var detailedErrorCode = -1;
  var storedException = null;

  try {
    var jsonSettings = JSON.stringify(serializedSettings);
    var settingsObjectToSave = {};
    settingsObjectToSave.SettingsKey = jsonSettings;
    saveSettingsRequest_OSF.DDA.ClientSettingsManager.write(settingsObjectToSave);
  } catch (ex) {
    storedException = ex;
  }

  var asyncResult = undefined;

  if (storedException != null) {
    detailedErrorCode = 9019;
    asyncResult = createAsyncResult(undefined, saveSettingsRequest_OSF.DDA.AsyncResultEnum.ErrorCode.Failed, detailedErrorCode, commonParameters.asyncContext, storedException.Message);
  } else {
    detailedErrorCode = 0;
    asyncResult = createAsyncResult(undefined, saveSettingsRequest_OSF.DDA.AsyncResultEnum.ErrorCode.Success, detailedErrorCode, commonParameters.asyncContext);
  }

  if (!!commonParameters.callback) {
    commonParameters.callback(asyncResult);
  }
}
// CONCATENATED MODULE: ./src/api/Settings.ts
var Settings_spreadArrays = undefined && undefined.__spreadArrays || function () {
  for (var s = 0, i = 0, il = arguments.length; i < il; i++) {
    s += arguments[i].length;
  }

  for (var r = Array(s), k = 0, i = 0; i < il; i++) {
    for (var a = arguments[i], j = 0, jl = a.length; j < jl; j++, k++) {
      r[k] = a[j];
    }
  }

  return r;
};



var Settings_OSF = __webpack_require__(0);

var Settings_Settings = function () {
  function Settings(deserializedData) {
    this.rawData = deserializedData;
    this.settingsData = null;
  }

  Settings.prototype.getSettingsData = function () {
    if (this.settingsData == null) {
      this.settingsData = this.convertFromRawSettings(this.rawData);
      this.rawData = null;
    }

    return this.settingsData;
  };

  Settings.prototype.get = function (key) {
    return this.getSettingsData()[key];
  };

  Settings.prototype.set = function (key, value) {
    this.getSettingsData()[key] = value;
  };

  Settings.prototype.remove = function (key) {
    delete this.getSettingsData()[key];
  };

  Settings.prototype.saveAsync = function () {
    var args = [];

    for (var _i = 0; _i < arguments.length; _i++) {
      args[_i] = arguments[_i];
    }

    saveSettingsRequest.apply(void 0, Settings_spreadArrays([this.getSettingsData()], args));
  };

  Settings.prototype.convertFromRawSettings = function (rawSettings) {
    if (rawSettings == null) {
      return {};
    }

    var outlookSettings = rawSettings.SettingsKey;

    if (!!outlookSettings) {
      return DeserializeRoamingSettings(outlookSettings);
    }

    return rawSettings;
  };

  return Settings;
}();


// CONCATENATED MODULE: ./src/api/Intellisense.ts



var Intellisense = {
  toItemRead: function toItemRead(item) {
    var hostItemtype = getHostItemType_getHostItemType();

    if (hostItemtype === HostItemType.Message || hostItemtype === HostItemType.Appointment || hostItemtype === HostItemType.MeetingRequest) {
      return item;
    }

    throw createArgumentTypeError();
  },
  toItemCompose: function toItemCompose(item) {
    var hostItemtype = getHostItemType_getHostItemType();

    if (hostItemtype === HostItemType.MessageCompose || hostItemtype === HostItemType.AppointmentCompose) {
      return item;
    }

    throw createArgumentTypeError();
  },
  toMessage: function toMessage(item) {
    return Intellisense.toMessageRead(item);
  },
  toMessageRead: function toMessageRead(item) {
    if (getHostItemType_getHostItemType() === HostItemType.Message || getHostItemType_getHostItemType() === HostItemType.MeetingRequest) {
      return item;
    }

    throw createArgumentTypeError();
  },
  toMessageCompose: function toMessageCompose(item) {
    if (getHostItemType_getHostItemType() === HostItemType.MessageCompose) {
      return item;
    }

    throw createArgumentTypeError();
  },
  toMeetingRequest: function toMeetingRequest(item) {
    if (getHostItemType_getHostItemType() === HostItemType.MeetingRequest) {
      return item;
    }

    throw createArgumentTypeError();
  },
  toAppointment: function toAppointment(item) {
    if (getHostItemType_getHostItemType() === HostItemType.Appointment) {
      return item;
    }

    throw createArgumentTypeError();
  },
  toAppointmentRead: function toAppointmentRead(item) {
    if (getHostItemType_getHostItemType() === HostItemType.Appointment) {
      return item;
    }

    throw createArgumentTypeError();
  },
  toAppointmentCompose: function toAppointmentCompose(item) {
    if (getHostItemType_getHostItemType() === HostItemType.AppointmentCompose) {
      return item;
    }

    throw createArgumentTypeError();
  }
};
// CONCATENATED MODULE: ./src/api/OutlookBase.ts


var OutlookBase = {
  SeriesTimeJsonConverter: function SeriesTimeJsonConverter(rawInput) {
    if (rawInput !== null && typeof rawInput === "object") {
      if (rawInput.seriesTimeJson !== null) {
        var seriesTime = new SeriesTime_SeriesTime();
        seriesTime.importFromSeriesTimeJsonObject(rawInput.seriesTimeJson);
        delete rawInput["seriesTimeJson"];
        rawInput.seriesTime = seriesTime;
      }
    }

    return rawInput;
  },
  CreateAttachmentDetails: function CreateAttachmentDetails(data) {
    convertAttachmentType(data);
    return data;
  }
};
// CONCATENATED MODULE: ./src/index.tsx






OSF = typeof OSF === "object" ? OSF : {};
OSF.DDA = OSF.DDA || {};
OSF.DDA.Settings = Settings_Settings;
OSF = typeof OSF === "object" ? OSF : {};
OSF.DDA = OSF.DDA || {};
OSF.DDA.OutlookAppOm = OutlookAppOm_OutlookAppOm;
Office = typeof Office === "object" ? Office : {};
Office.cast = Office.cast || {};
Office.cast.item = Intellisense;
Microsoft.Office.WebExtension.MailboxEnums = MailboxEnums;
Microsoft.Office.WebExtension.CoercionType = CoercionType;
Microsoft.Office.WebExtension.SeriesTime = SeriesTime_SeriesTime;
Microsoft.Office.WebExtension.OutlookBase = OutlookBase;
/* harmony default export */ var src = __webpack_exports__["default"] = (OutlookAppOm_OutlookAppOm);

/***/ })
/******/ ])["default"];
//# sourceMappingURL=outlook.launchEvent.js.map
    OSF.DDA.ErrorCodeManager.initializeErrorMessages(Strings.OfficeOM);
    // Office.context.mailbox = new OSF.DDA.OutlookAppOm(appContext, this._webAppState.wnd, appReady);
    Office.context.mailbox = new OSF.DDA.OutlookAppOm(appContext, {}, appReady);
    InitRoamingSettings();
    CustomFunctionMappings['SendOmReadyMessageToHost']();
};

let GlobalMailboxInitialData;

function OnMailboxOMReady()
{
}

window = {};

function MailboxCreateOM(initialData)
{
    GlobalMailboxInitialData = JSON.parse(initialData);
    MailboxCreateOMHelper(OSF.OfficeAppContext, OnMailboxOMReady, window);
}

function InitRoamingSettings()
{
    let deserializedData;
    let roamingSettingsData = GlobalMailboxInitialData.roamingSettings;

    if (roamingSettingsData)
    {
        deserializedData = DeserializeRoamingSettings({ SettingsKey : roamingSettingsData });
    }

    if (deserializedData)
    {
      Office.context.roamingSettings = new OSF.DDA.Settings(deserializedData);
    }
    else
    {
        console.log("Failed to initialize Roaming Settings. 'roamingSettings' NOT found in initial data blob.");
    }
}

function DeserializeRoamingSettings(serializedSettings)
{
    let ret = {};
    serializedSettings = serializedSettings || {};

    const DateJSONPrefix = "Date(";
    const DataJSONSuffix = ")";

    for (let key in serializedSettings)
    {
        let value = serializedSettings[key];

        try
        {
            value = JSON.parse(
                value,
                function dateReviver(k, v)
                {
                    let d;
                    if (typeof v === 'string' && v && v.length > 6 && v.slice(0, 5) === DateJSONPrefix && v.slice(-1) === DataJSONSuffix)
                    {
                        d = new Date(parseInt(v.slice(5, -1)));
                        if (d)
                        {
                            return d;
                        }
                    }

                    return v;
                }
            );

            ret[key] = value;
        }
        catch(ex)
        {
        }
    }

    return ret;
}
// ================================================================================================================================================
// OUTLOOKJS CODE ENDS HERE
// ================================================================================================================================================