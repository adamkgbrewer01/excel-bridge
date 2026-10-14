"""Write manifest.xml, the file Excel reads to find the add-in and add its ribbon button.

Run:  python3 tools/make_manifest.py https://example.github.io/excel-bridge
The address is wherever taskpane.html is hosted. It must be https.
"""
import sys
from pathlib import Path
from urllib.parse import urlparse
from xml.dom import minidom

ADDIN = Path(__file__).resolve().parent.parent
ADDIN_ID = "7d6dabaf-34fa-4540-b4aa-f98a073c9248"
VERSION = "1.0.0.0"

TEMPLATE = """<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
           xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
           xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
           xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides"
           xsi:type="TaskPaneApp">
  <Id>{id}</Id>
  <Version>{version}</Version>
  <ProviderName>Excel Bridge</ProviderName>
  <DefaultLocale>en-GB</DefaultLocale>
  <DisplayName DefaultValue="Excel Bridge"/>
  <Description DefaultValue="Dashboards in Excel where charts filter each other, drill down and split into small multiples."/>
  <IconUrl DefaultValue="{base}/assets/icon-32.png"/>
  <HighResolutionIconUrl DefaultValue="{base}/assets/icon-64.png"/>
  <SupportUrl DefaultValue="{base}/taskpane.html"/>
  <AppDomains>
    <AppDomain>{origin}</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Workbook"/>
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="ExcelApi" MinVersion="1.7"/>
    </Sets>
  </Requirements>
  <DefaultSettings>
    <SourceLocation DefaultValue="{base}/taskpane.html"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Workbook">
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="GetStarted.Title"/>
            <Description resid="GetStarted.Description"/>
            <LearnMoreUrl resid="Taskpane.Url"/>
          </GetStarted>
          <FunctionFile resid="Commands.Url"/>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="ExcelBridge.Group">
                <Label resid="Group.Label"/>
                <Icon>
                  <bt:Image size="16" resid="Icon.16"/>
                  <bt:Image size="32" resid="Icon.32"/>
                  <bt:Image size="80" resid="Icon.80"/>
                </Icon>
                <Control xsi:type="Button" id="ExcelBridge.Open">
                  <Label resid="Button.Label"/>
                  <Supertip>
                    <Title resid="Button.Label"/>
                    <Description resid="Button.Tooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16"/>
                    <bt:Image size="32" resid="Icon.32"/>
                    <bt:Image size="80" resid="Icon.80"/>
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>ExcelBridgePane</TaskpaneId>
                    <SourceLocation resid="Taskpane.Url"/>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16" DefaultValue="{base}/assets/icon-16.png"/>
        <bt:Image id="Icon.32" DefaultValue="{base}/assets/icon-32.png"/>
        <bt:Image id="Icon.80" DefaultValue="{base}/assets/icon-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Commands.Url" DefaultValue="{base}/commands.html"/>
        <bt:Url id="Taskpane.Url" DefaultValue="{base}/taskpane.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="GetStarted.Title" DefaultValue="Excel Bridge is ready"/>
        <bt:String id="Group.Label" DefaultValue="Excel Bridge"/>
        <bt:String id="Button.Label" DefaultValue="Dashboard"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="GetStarted.Description" DefaultValue="Click Dashboard on the Home tab to open it."/>
        <bt:String id="Button.Tooltip" DefaultValue="Open a dashboard for a table, with charts that filter each other and drill down."/>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
"""


def main():
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    base = sys.argv[1].rstrip("/")
    url = urlparse(base)
    if url.scheme != "https":
        sys.exit("The address must start with https://")
    xml = TEMPLATE.format(id=ADDIN_ID, version=VERSION, base=base, origin=f"{url.scheme}://{url.netloc}")
    minidom.parseString(xml.encode())  # fails loudly if the XML is malformed
    (ADDIN / "manifest.xml").write_text(xml)
    print("manifest.xml written for", base)


if __name__ == "__main__":
    main()
