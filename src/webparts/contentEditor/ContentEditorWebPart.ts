import { Version } from '@microsoft/sp-core-library';
import { SPComponentLoader } from '@microsoft/sp-loader';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './ContentEditor.module.scss';
import * as strings from 'contentEditorStrings';
import { IContentEditorWebPartProps } from './IContentEditorWebPartProps';

import {
  SPHttpClient
} from '@microsoft/sp-http';

import {
  Environment,
  EnvironmentType
} from '@microsoft/sp-core-library';

export default class ContentEditorWebPart extends BaseClientSideWebPart<IContentEditorWebPartProps> {

  public render(): void {
    if (Environment.type == EnvironmentType.Local) {
      this.domElement.innerHTML = `
      <div id="weather"></div>
      <script src="//code.jquery.com/jquery-2.1.1.min.js"></script>
      <script src="//cdnjs.cloudflare.com/ajax/libs/jquery.simpleWeather/3.1.0/jquery.simpleWeather.min.js"></script>
      <script>
        jQuery.simpleWeather({
        location: 'Oslo, Norway',
        woeid: '',
        unit: 'c',
        success: function(weather) {
          html = '<h2>'+weather.temp+'&deg;'+weather.units.temp+'</h2>';
          html += '<ul><li>'+weather.city+', '+weather.region+'</li>';
          html += '<li>'+weather.currently+'</li></ul>';
        
          jQuery("#weather").html(html);
        },
        error: function(error) {
          jQuery("#weather").html('<p>'+error+'</p>');
        }
        });
        </script>`;
        this.executeScript(this.domElement);
    } else {
      if (this.properties.contentLink) {
        // Add _spPageContextInfo global variable 
        let w = (window as any);
        if (!w._spPageContextInfo) {
          w._spPageContextInfo = this.context.pageContext.legacyPageContext;
        }

        // Add form digest hidden field
        if (!document.getElementById('__REQUESTDIGEST')) {
          let digestValue = this.context.pageContext.legacyPageContext.formDigestValue;
          let requestDigestInput: Element = document.createElement('input');
          requestDigestInput.setAttribute('type', 'hidden');
          requestDigestInput.setAttribute('name', '__REQUESTDIGEST');
          requestDigestInput.setAttribute('id', '__REQUESTDIGEST');
          requestDigestInput.setAttribute('value', digestValue);
          document.body.appendChild(requestDigestInput);        
        }

        // Get server relative URL to content link file
        let filePath = this.properties.contentLink;
        if (filePath.toLowerCase().substr(0, 4) == "http") {
          let parts = filePath.replace("://", "").split("/");
          parts.shift();
          filePath = "/" + parts.join("/");
        }

        // Get file and read script
        let siteUrl = this.context.pageContext.site.absoluteUrl;
        this.context.spHttpClient.get(siteUrl + 
          "/_api/Web/getFileByServerRelativeUrl('" + filePath + "')/$value", 
          SPHttpClient.configurations.v1)
          .then((response) => {
            return response.text();
          })
          .then((value) => {
            this.domElement.innerHTML = ``;

            // Ensure the Client Object Model is loaded
            if (!w.SP) {
              this.domElement.innerHTML += `
              <script type="text/javascript" src="${siteUrl}/_layouts/15/init.js"></script>
              <script type="text/javascript" src="${siteUrl}/_layouts/15/MicrosoftAjax.js"></script>
              <script type="text/javascript" src="${siteUrl}/_layouts/15/SP.Runtime.js"></script>
              <script type="text/javascript" src="${siteUrl}/_layouts/15/SP.js"></script>
              `;  
            }

            this.domElement.innerHTML += value;
            this.executeScript(this.domElement);
          });
      }
    }    
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('contentLink', {
                  label: strings.ContentLinkFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }

  // Finds and executes scripts in a newly added element's body.
  // Needed since innerHTML does not run scripts.
  //
  // Argument element is an element in the dom.
  private executeScript(element: HTMLElement) {
    // Define global name to tack scripts on in case script to be loaded is not AMD/UMD
    (<any>window).ScriptGlobal = {};

    function nodeName(elem, name) {
      return elem.nodeName && elem.nodeName.toUpperCase() === name.toUpperCase();
    }

    function evalScript(elem) {
      var data = (elem.text || elem.textContent || elem.innerHTML || ""),
        head = document.getElementsByTagName("head")[0] ||
          document.documentElement,
        script = document.createElement("script");

      script.type = "text/javascript";
      if (elem.src && elem.src.length > 0) {
        return;
      }
      if (elem.onload && elem.onload.length > 0) {
        script.onload = elem.onload;
      }

      try {
        // doesn't work on ie...
        script.appendChild(document.createTextNode(data));
      } catch (e) {
        // IE has funky script nodes
        script.text = data;
      }

      head.insertBefore(script, head.firstChild);
      head.removeChild(script);
    }

    // main section of function
    var scripts = [],
      script,
      children_nodes = element.childNodes,
      child,
      i;

    for (i = 0; children_nodes[i]; i++) {
      child = children_nodes[i];
      if (nodeName(child, "script") &&
        (!child.type || child.type.toLowerCase() === "text/javascript")) {
        scripts.push(child);
      }
    }

    const urls = [];
    const onLoads = [];
    for (i = 0; scripts[i]; i++) {
      script = scripts[i];
      if (script.src && script.src.length > 0) {
        urls.push(script.src);
      }
      if (script.onload && script.onload.length > 0) {
        onLoads.push(script.onload);
      }
    }

    // Execute promises in sequentially - https://hackernoon.com/functional-javascript-resolving-promises-sequentially-7aac18c4431e
    // Use "ScriptGlobal" as the global namein case script is AMD/UMD
    const allFuncs = urls.map(url => () => SPComponentLoader.loadScript(url, { globalExportsName: "ScriptGlobal" }));

    const promiseSerial = funcs =>
      funcs.reduce((promise, func) =>
        promise.then(result => func().then(Array.prototype.concat.bind(result))),
        Promise.resolve([]));

    // execute Promises in serial
    promiseSerial(allFuncs)
      .then(() => {
        // execute any onload people have added
        for (i = 0; onLoads[i]; i++) {
          onLoads[i]();
        }
        // execute script blocks
        for (i = 0; scripts[i]; i++) {
          script = scripts[i];
          if (script.parentNode) { script.parentNode.removeChild(script); }
          evalScript(scripts[i]);
        }
      }).catch(console.error);
  };  

}
