Office.onReady(() => {
  // Add your code here to initialize your extension

  // Button click event handler
  function onButtonClick(event) {
    // Add your code here to handle the button click
    console.log("Button clicked!");
  }

  // Register the button click event handler
  Office.ribbon.requestUpdate({
    controls: [
      {
        id: "customButton",
        onAction: onButtonClick
      }
    ]
  });
});

