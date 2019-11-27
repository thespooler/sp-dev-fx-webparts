import { WebPartContext } from "@microsoft/sp-webpart-base";
import { sp } from '@pnp/sp';

export default class spservices {

  constructor(private context:WebPartContext) {

    sp.setup({
      spfxContext: this.context

    });
  }

  public async getUserProperties(user:string){

    let currentUserProperties:any = await sp.profiles.getPropertiesFor(user);
    console.log(currentUserProperties);

    return currentUserProperties;
  }

  public async getUserProfileProperty(user:string,property:string) {
    let UserProperty:any = await sp.profiles.getUserProfilePropertyFor(user, property);
    console.log(UserProperty);

    return UserProperty;
  }   
}
