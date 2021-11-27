import * as ews from "ews-javascript-api";
import * as util from "util";

//create AutodiscoverService object
const autod = new ews.AutodiscoverService(ews.ExchangeVersion.Exchange2013);

//you can omit url and it will autodiscover the url, version helps throw error on client side for unsupported operations.example - //var autod = new ews.AutodiscoverService(ews.ExchangeVersion.Exchange2010);
//set credential for service
autod.Credentials = new ews.WebCredentials("corp\\madridme", "Vmt8Pb7QpTF2XHN");

//create array to include list of desired settings
const settings: ews.UserSettingName[] = [
  ews.UserSettingName.InternalEwsUrl,
  ews.UserSettingName.ExternalEwsUrl,
  ews.UserSettingName.UserDisplayName,
  ews.UserSettingName.UserDN,
  ews.UserSettingName.EwsPartnerUrl,
  ews.UserSettingName.DocumentSharingLocations,
  ews.UserSettingName.MailboxDN,
  ews.UserSettingName.ActiveDirectoryServer,
  ews.UserSettingName.CasVersion,
  ews.UserSettingName.ExternalWebClientUrls,
  ews.UserSettingName.ExternalImap4Connections,
  ews.UserSettingName.AlternateMailboxes,
];
//get the setting

autod.GetUserSettings("enrique.madrid@italtel.com", settings).then(
  (response) => {
    // console.log(autod.Url.ToString());
    // uncoment next line to see full response from autodiscover, you will need to add var util = require('util');
    // console.log(
    //   util.inspect(response, { showHidden: false, depth: null, colors: true })
    // );
    response.Settings.Items.forEach((value) => {
      console.log(
        ews.StringHelper.Format(
          "{0} = {1}",
          ews.UserSettingName[value.key],
          value.value
        )
      );
    });
  },
  (e) => {
    //log errors or do something with errors
    console.log("Error " + e);
  }
);
