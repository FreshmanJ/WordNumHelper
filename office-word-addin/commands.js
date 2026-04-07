Office.onReady(() => {});

function noop(event) {
  event.completed();
}

if (typeof Office !== "undefined" && Office.actions) {
  Office.actions.associate("noop", noop);
}
