import { EwsService } from "./EwsService";
import { IExchangeService } from "./IExchangeService";
import { RestService } from "./RestService";

export class ExchangeServiceFactory {

    private static minRestSet: string = "1.5";
  
    private static _service : IExchangeService;
  
    public static service() : IExchangeService {
      if (!ExchangeServiceFactory._service) {
        ExchangeServiceFactory._service = ExchangeServiceFactory.createExchangeService();
      }
  
      return ExchangeServiceFactory._service;
    }
  
    private static canUseAPI(apiType: string, minset: string): boolean {
      if (typeof (Office) === "undefined") { return false; }
      if (!Office) { return false; }
      if (!Office.context) { return false; }
      if (!Office.context.requirements) { return false; }
      if (!Office.context.requirements.isSetSupported("Mailbox", minset)) { return false; }
      if (!Office.context.mailbox) { return false; }
      if (!Office.context.mailbox.getCallbackTokenAsync) { return false; }
      return true;
    }
  
    private static canUseRest(): boolean { return ExchangeServiceFactory.canUseAPI("Rest", ExchangeServiceFactory.minRestSet); }
  
    private static createExchangeService() : IExchangeService {
      if (ExchangeServiceFactory.canUseRest()) {
        return new RestService();
      } else {
        return new EwsService();
      }
    }
  }