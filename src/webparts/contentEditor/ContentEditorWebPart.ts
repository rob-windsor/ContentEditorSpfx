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

  private lastContentLink: string = null;

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
      debugger;
      if (this.properties.ContentLink == null || this.properties.ContentLink.trim().length == 0) {
        this.domElement.innerHTML = "You have not entered a value for the Content Link property.";
      } else if (this.lastContentLink != null) {
        if (this.lastContentLink != this.properties.ContentLink) {
          window.setTimeout(() => {
            alert("You have changed the value of the Content Link property. The page needs to reload.");
            window.location.reload(true);
          },1000);  
        }
      } else {
        this.lastContentLink = this.properties.ContentLink;
        
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
        let filePath = this.properties.ContentLink;
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

  private evalScript(elem) {
    const data = (elem.text || elem.textContent || elem.innerHTML || "");
    const headTag = document.getElementsByTagName("head")[0] || document.documentElement;
    const scriptTag = document.createElement("script");

    scriptTag.type = "text/javascript";
    if (elem.src && elem.src.length > 0) {
      return;
    }
    if (elem.onload && elem.onload.length > 0) {
      scriptTag.onload = elem.onload;
    }

    try {
      // doesn't work on ie...
      scriptTag.appendChild(document.createTextNode(data));
    } catch (e) {
      // IE has funky script nodes
      scriptTag.text = data;
    }

    //console.log(scriptTag.innerHTML);
    headTag.insertBefore(scriptTag, headTag.firstChild);
    headTag.removeChild(scriptTag);
  }

  private nodeName(elem, name) {
    return elem.nodeName && elem.nodeName.toUpperCase() === name.toUpperCase();
  }

  // Finds and executes scripts in a newly added element's body.
  // Needed since innerHTML does not run scripts.
  //
  // Argument element is an element in the dom.
  private async executeScript(element: HTMLElement) {
    // Define global name to tack scripts on in case script to be loaded is not AMD/UMD
    (<any>window).ScriptGlobal = {};

    // main section of function
    const scripts = [];
    const children_nodes = element.childNodes;

    for (let i = 0; children_nodes[i]; i++) {
      const child: any = children_nodes[i];
      if (this.nodeName(child, "script") &&
        (!child.type || child.type.toLowerCase() === "text/javascript")) {
        scripts.push(child);
      }
    }

    const urls = [];
    const onLoads = [];
    for (let i = 0; scripts[i]; i++) {
      const scriptTag = scripts[i];
      if (scriptTag.src && scriptTag.src.length > 0) {
        urls.push(scriptTag.src);
      }
      if (scriptTag.onload && scriptTag.onload.length > 0) {
        onLoads.push(scriptTag.onload);
      }
    }

    let oldamd = null;
    if (window["define"] && window["define"].amd) {
      oldamd = window["define"].amd;
      window["define"].amd = null;
    }

    for (let i = 0; i < urls.length; i++) {
      try {
        await SPComponentLoader.loadScript(urls[i], { globalExportsName: "ScriptGlobal" });
      } catch (error) {
        console.error(error);
      }
    }
    if (oldamd) {
      window["define"].amd = oldamd;
    }

    for (let i = 0; scripts[i]; i++) {
      const scriptTag = scripts[i];
      if (scriptTag.parentNode) { scriptTag.parentNode.removeChild(scriptTag); }
      this.evalScript(scripts[i]);
    }
    // execute any onload people have added
    for (let i = 0; onLoads[i]; i++) {
      onLoads[i]();
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
                PropertyPaneTextField('ContentLink', {
                  label: strings.ContentLinkFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
