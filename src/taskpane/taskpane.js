// The initialize function must be run each time a new page is loaded.
Office.initialize = () => {
  // Show UI only after everything is initialized.
  document.getElementById('app-body').style.display = 'flex';
};
