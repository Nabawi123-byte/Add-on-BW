// Membuat manifest-excel.xml dan manifest-word.xml.
// Pakai: node scripts/manifest.js [--port 3000]

import { writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import path from 'node:path';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const portArg = process.argv.indexOf('--port');
const PORT = portArg > 0 ? Number(process.argv[portArg + 1]) : 3000;
const BASE = `https://localhost:${PORT}`;

const HOSTS = {
  excel: { id: '6bca5ac5-67dd-4fe4-a4a2-499b1f211663', host: 'Workbook', app: 'Excel' },
  word: { id: '9ffd2c0a-6c65-4c46-af3e-61672104b98f', host: 'Document', app: 'Word' },
};

function manifest({ id, host, app }) {
  const functions =
    host === 'Workbook'
      ? `
        <AllFormFactors>
          <ExtensionPoint xsi:type="CustomFunctions">
            <Script>
              <SourceLocation resid="Functions.Script.Url"/>
            </Script>
            <Page>
              <SourceLocation resid="Functions.Page.Url"/>
            </Page>
            <Metadata>
              <SourceLocation resid="Functions.Metadata.Url"/>
            </Metadata>
            <Namespace resid="Functions.Namespace"/>
          </ExtensionPoint>
        </AllFormFactors>`
      : '';
  const fnUrls =
    host === 'Workbook'
      ? `
        <bt:Url id="Functions.Script.Url" DefaultValue="${BASE}/functions.js"/>
        <bt:Url id="Functions.Metadata.Url" DefaultValue="${BASE}/functions.json"/>
        <bt:Url id="Functions.Page.Url" DefaultValue="${BASE}/functions.html"/>`
      : '';
  const fnStrings = host === 'Workbook' ? `\n        <bt:String id="Functions.Namespace" DefaultValue="MASTER"/>` : '';
  const icons = `
                  <Icon>
                    <bt:Image size="16" resid="Icon.16"/>
                    <bt:Image size="32" resid="Icon.32"/>
                    <bt:Image size="80" resid="Icon.80"/>
                  </Icon>`;
  return `<?xml version="1.0" encoding="UTF-8"?>
<!-- Dibuat oleh scripts/manifest.js — jangan edit manual; jalankan: npm run manifest -->
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
           xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
           xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
           xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides"
           xsi:type="TaskPaneApp">
  <Id>${id}</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Master</ProviderName>
  <DefaultLocale>id-ID</DefaultLocale>
  <DisplayName DefaultValue="Master"/>
  <Description DefaultValue="Asisten perintah untuk ${app}: template keuangan, surat otomatis, tabel, dan AI Claude. Perintah bawaan berjalan offline."/>
  <IconUrl DefaultValue="${BASE}/assets/icon-32.png"/>
  <HighResolutionIconUrl DefaultValue="${BASE}/assets/icon-64.png"/>
  <SupportUrl DefaultValue="https://github.com/nabawi123-byte/add-on-bw"/>
  <AppDomains>
    <AppDomain>https://api.anthropic.com</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="${host}"/>
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="${BASE}/taskpane.html"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="${host}">${functions}
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="GetStarted.Title"/>
            <Description resid="GetStarted.Description"/>
            <LearnMoreUrl resid="GetStarted.LearnMoreUrl"/>
          </GetStarted>
          <FunctionFile resid="Commands.Url"/>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="Master.Group">
                <Label resid="Group.Label"/>${icons.replace(/^ {18}/gm, '                ')}
                <Control xsi:type="Button" id="Master.Open">
                  <Label resid="Button.Label"/>
                  <Supertip>
                    <Title resid="Button.Label"/>
                    <Description resid="Button.Tooltip"/>
                  </Supertip>${icons}
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>MasterPane</TaskpaneId>
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
        <bt:Image id="Icon.16" DefaultValue="${BASE}/assets/icon-16.png"/>
        <bt:Image id="Icon.32" DefaultValue="${BASE}/assets/icon-32.png"/>
        <bt:Image id="Icon.80" DefaultValue="${BASE}/assets/icon-80.png"/>
      </bt:Images>
      <bt:Urls>${fnUrls}
        <bt:Url id="GetStarted.LearnMoreUrl" DefaultValue="https://github.com/nabawi123-byte/add-on-bw"/>
        <bt:Url id="Commands.Url" DefaultValue="${BASE}/commands.html"/>
        <bt:Url id="Taskpane.Url" DefaultValue="${BASE}/taskpane.html"/>
      </bt:Urls>
      <bt:ShortStrings>${fnStrings}
        <bt:String id="GetStarted.Title" DefaultValue="Master siap dipakai"/>
        <bt:String id="Group.Label" DefaultValue="Master"/>
        <bt:String id="Button.Label" DefaultValue="Master"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="GetStarted.Description" DefaultValue="Klik tombol Master di tab Beranda (Home), lalu ketik perintah, misalnya 'buat anggaran bulanan' atau 'buat surat izin sakit'."/>
        <bt:String id="Button.Tooltip" DefaultValue="Buka panel Master untuk menjalankan perintah, template, dan AI."/>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
`;
}

for (const [name, cfg] of Object.entries(HOSTS)) {
  const file = path.join(root, `manifest-${name}.xml`);
  writeFileSync(file, manifest(cfg));
  console.log('ditulis', path.relative(root, file));
}
