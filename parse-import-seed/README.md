*** parseImportSeed javascript function***

This function is designed to
- accept an Excel file as an arrayBuffer in a client-side file upload event handler
- validate the contents of the Excel file
- return the parsed contents of the Excel file as a javascript object

**How to use this method :**

First, you need to import it as a script into your application client code :

```
  <script src="https://nabu-io.github.io/public-js/parseImportSeed.js"></script>
```

Second, include the parser in the event handler for the Excel file upload :

```
  const fileReader = new FileReader();

  fileReader.onload = async (event) => {
    const arrayBuffer = event.target.result;

    parseImportSeed(arrayBuffer)
      .then((parsedData) => {
        // send the data to the API as JSON
      })
      .catch((error) => {
        // handle parsing error - show error to client
      })
  }

  fileReader.onerror = (event) => {
      console.error("File could not be read: " + fileReader.error);
  };

  fileReader.readAsArrayBuffer(file);

```

**To run the tests**

```
  npm run test
```