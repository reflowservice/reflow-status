/* REFLOW STATUS — strato cloud (sola lettura).
   Login Microsoft (MSAL, registrazione condivisa "REFLOW ONE" con PLANNER, iMAPS e ORBIT)
   e lettura via Microsoft Graph del REFLOW MASTER direttamente dalla libreria "Reflow - Operativo".
   L'app non scrive nulla su SharePoint. */
(function () {
  'use strict';
  const CLIENT_ID = 'ef16e670-e362-4954-a912-b628afde91a0';
  const TENANT_ID = 'd53a79c7-2449-4eca-be76-4bd4b7482c96';
  const SITE_HOST = 'reflowservice.sharepoint.com';
  const SITE_PATH = '/sites/ReflowArchive';
  const MASTER = '3_CONTROL ROOM/REFLOW MASTER.xlsx';
  const SCOPES = ['User.Read', 'Sites.ReadWrite.All', 'Files.ReadWrite.All'];
  const PAGES_URL = 'https://reflowservice.github.io/reflow-status/';

  const redirectUri = location.hostname.endsWith('github.io') ? PAGES_URL : location.origin + location.pathname;
  const pca = new msal.PublicClientApplication({
    auth: { clientId: CLIENT_ID, authority: `https://login.microsoftonline.com/${TENANT_ID}`, redirectUri },
    cache: { cacheLocation: 'localStorage' }
  });

  const S = window.STATUS = { account: null, ready: null };
  S.login = () => pca.loginRedirect({ scopes: SCOPES, prompt: 'select_account' });
  S.logout = () => pca.logoutRedirect({ account: S.account, postLogoutRedirectUri: redirectUri });

  async function token() {
    try {
      return (await pca.acquireTokenSilent({ scopes: SCOPES, account: S.account })).accessToken;
    } catch (e) {
      if (e instanceof msal.InteractionRequiredAuthError) { await pca.acquireTokenRedirect({ scopes: SCOPES, account: S.account }); }
      throw e;
    }
  }
  async function graph(path, opt = {}) {
    const t = await token();
    const headers = Object.assign({ Authorization: 'Bearer ' + t }, opt.headers || {});
    return fetch('https://graph.microsoft.com/v1.0' + path, Object.assign({}, opt, { headers }));
  }
  let siteId = null;
  async function site() {
    if (siteId) return siteId;
    const r = await graph(`/sites/${SITE_HOST}:${SITE_PATH}?$select=id`);
    if (!r.ok) throw new Error('Sito SharePoint non raggiungibile (' + r.status + ')');
    return siteId = (await r.json()).id;
  }
  const item = rel => `/sites/${siteId}/drive/root:/${encodeURI(rel).replace(/#/g, '%23')}:`;

  /* Ritorna { data: ArrayBuffer, modified: ISO, by: nome, url: link al file } */
  S.loadMaster = async () => {
    await site();
    const m = await graph(item(MASTER) + '?$select=lastModifiedDateTime,lastModifiedBy,webUrl', { cache: 'no-store' });
    if (m.status === 404) throw new Error('REFLOW MASTER.xlsx non trovato in 3_CONTROL ROOM.');
    if (!m.ok) throw new Error('Lettura del Master non riuscita (' + m.status + ')');
    const meta = await m.json();
    const r = await graph(item(MASTER) + '/content', { cache: 'no-store' });
    if (!r.ok) throw new Error('Download del Master non riuscito (' + r.status + ')');
    return {
      data: await r.arrayBuffer(),
      modified: meta.lastModifiedDateTime,
      by: meta.lastModifiedBy && meta.lastModifiedBy.user ? meta.lastModifiedBy.user.displayName : '',
      url: meta.webUrl || ''
    };
  };

  S.ready = (async () => {
    await pca.initialize();
    const r = await pca.handleRedirectPromise();
    S.account = (r && r.account) || pca.getAllAccounts()[0] || null;
    return !!S.account;
  })();
})();
