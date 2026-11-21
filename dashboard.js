(() => {
  const DataBook = window.DataBook || (window.DataBook = {});
  DataBook.dashboard = {
    init(showView) {
      document.getElementById('open-order-summary')
        .addEventListener('click', () => showView('order-summary'));
      document.getElementById('open-data')
        .addEventListener('click', () => showView('data'));
    }
  };
})();
