import { spfi, SPFI, SPFx } from "@pnp/sp";
import { LogLevel, PnPLogging } from "@pnp/logging";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/batching";
import { WebPartContext } from "@microsoft/sp-webpart-base/dist/index-internal";

let sp: SPFI;

export const getSP = (context?: WebPartContext): SPFI => {
  if (context !== null && context !== undefined) {
    sp = spfi().using(SPFx(context)).using(PnPLogging(LogLevel.Warning));
  }
  return sp;
};
