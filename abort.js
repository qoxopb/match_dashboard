let abortRequested = false;
let activePage = null;

module.exports = {
  shouldAbort: () => abortRequested,
  requestAbort: () => {
    abortRequested = true;
    if (activePage) activePage.close().catch(() => {});
  },
  resetAbort: () => { abortRequested = false; activePage = null; },
  setActivePage: (page) => { activePage = page; },
};
