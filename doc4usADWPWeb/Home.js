(function () {
    "use strict";

    var messageBanner;
    var selId;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the notification mechanism and hide it
            var element = document.querySelector('.MessageBanner');
            messageBanner = new components.MessageBanner(element);
            messageBanner.hideBanner();

            // If not using Word 2016, use fallback logic.
            if (!Office.context.requirements.isSetSupported('WordApi', '1.1')) {
                $("#template-description").text("Este ejemplo muestra el texto sleccionado.");
                $('#button-text').text("Mostrar!");
                $('#button-desc').text("Mostrar texto seleccionado");

                $('#highlight-button').click(displaySelectedText);
                return;
            }
            else {
                // Just letting you know that this code will not work with your version of Word.
                $('#supportedVersion').html('This code requires Word 2016 or later.');
            }


            $("#template-description").text("Este ejemplo ilumina un trozo de texto dentro del documento para probar las funcionalidades del add-ins requeridas.");
            $('#button-text').text("Boton de Prueba de Mensajeria");
            $('#button-desc').text("Texto largo de  demostración");


            // Eventos Click de los botones de Plantillas o Documentos
            $('#gestionaplantilla').click(GestionarPlantilla);
            $('#gestionadocumento').click(GestionarDocumento);
            $('#VerDocumento').click(VerDocumento);
            $('#GuardarDocumento').click(GuardarDocumento);
            $('#VerPreliminar').click(VerPreliminar);
            $('#guardarPlantilla').click(guardarPlantilla);

            $('#combinarretiqueta').click(combineText);
            $('#insertaretiqueta').click(insertarEtiqueta);
            $('#TipoEtiqueta').change(getEtiquetas);
            $('#supportedVersion').html('Esta Add-ins de Doc4us debe ser usado para Word 2016 o Superior.');
            getClasificacionEtiquetas();

            loadSampleData();
            // Add a click event handler for the highlight button.
            $('#highlight-button').click(hightlightLongestWord);
        });
    };




    // Get the document mode and the URL of the active project.
    function showDocumentProperties() {
        var output = String.format(
            'The document mode is {0}.<br/>The URL of the active project is {1}.',
            Office.context.document.mode,
            Office.context.document.url);
        $('#message').html(output);
    }

    function guardarPlantilla() {

        //Get the URL of the current file.
        Office.context.document.getFilePropertiesAsync(function (asyncResult) {
            var fileUrl = asyncResult.value.url;
            if (fileUrl === "") {
                showNotification('Archivo', 'archivo no guardado, debe guardar documento ');
            }
            else {
                //showNotification('Archivo', fileUrl);
                //var selValue = Office.context.document.url;
                let value = fileUrl.split('/').reverse()[0];
                let fileName = value.split('.')[0];
                let _tipoDocumento = fileName.split("-")[1];
                let _Id = fileName.split("-")[2];


                var selApp = document.getElementById("App");
                selId = document.getElementById("IdDocumento");
                selApp.value = _tipoDocumento;
                selId.value = _Id;

                Word.run(function (context) {

                    var thisDocument = context.document;

                    var range = thisDocument.getSelection();
                    range.insertText(fileUrl);

                    //_UrlFile = Office.context.document.url;
                    context.load(thisDocument, 'saved');
                    return context.sync().then(function () {
                        if (thisDocument.saved === false) {
                            thisDocument.save();
                            getFileAsyncInternal();
                        } else {
                            getFileAsyncInternal();
                        }
                    });
                })
                    .catch(function (error) {
                        showNotification("Error al Guardar", JSON.stringify(error));

                        console.log("Error4: " + JSON.stringify(error));
                        if (error instanceof OfficeExtension.Error) {
                            console.log("Debug info: " + JSON.stringify(error.debugInfo));
                        }
                    });

            }
            return fileUrl;
        });
    }

    function getFileAsyncInternal() {
        Office.context.document.getFileAsync("compressed", { sliceSize: 10240 }, function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                document.getElementById("log").textContent = JSON.stringify(asyncResult);
            }
            else {
                getAllSlices(asyncResult.value);
            }
        });
    }




    // Get all the slices of file from the host after "getFileAsync" is done.
    function getAllSlices(file) {
        var sliceCount = file.sliceCount;
        var sliceIndex = 0;
        var docdata = [];
        var getSlice = function () {
            file.getSliceAsync(sliceIndex, function (asyncResult) {
                if (asyncResult.status == "succeeded") {
                    docdata = docdata.concat(asyncResult.value.data);
                    sliceIndex++;
                    if (sliceIndex == sliceCount) {
                        file.closeAsync();
                        onGetAllSlicesSucceeded(docdata);
                    }
                    else {
                        getSlice();
                    }
                }
                else {
                    file.closeAsync();
                    document.getElementById("log").textContent = JSON.stringify(asyncResult);
                }
            });
        };
        getSlice();
    }




    function myEncodeBase64(docData) {

        var binary = '';
        var bytes = new Uint8Array(docData);
        var len = bytes.byteLength;
        for (var i = 0; i < len; i++) {
            binary += String.fromCharCode(bytes[i]);
        }
        return window.btoa(binary);
    }

    function sendSlice(slice, state) {
        var data = slice.data;

        // If the slice contains data, create an HTTP request.
        if (data) {

            // Encode the slice data, a byte array, as a Base64 string.
            // NOTE: The implementation of myEncodeBase64(input) function isn't
            // included with this example. For information about Base64 encoding with
            // JavaScript, see https://developer.mozilla.org/docs/Web/JavaScript/Base64_encoding_and_decoding.
            var fileData = myEncodeBase64(data);
            // Create a new HTTP request. You need to send the request
            // to a webpage that can receive a post.
            var request = new XMLHttpRequest();

            // Create a handler function to update the status
            // when the request has been sent.
            request.onreadystatechange = function () {
                if (request.readyState == 4) {
                    state.counter++;
                    if (state.counter < state.sliceCount) {
                        getSlice(state);
                    }
                    else {
                        closeFile(state);
                    }
                }
            }

            request.open("POST", "[Your receiving page or service]");
            request.setRequestHeader("Slice-Number", slice.index);

            // Send the file as the body of an HTTP POST
            // request to the web server.
            request.send(fileData);
        }
    }

    function closeFile(state) {
        // Close the file when you're done with it.
        state.file.closeAsync(function (result) {

            // If the result returns as a success, the
            // file has been successfully closed.
            if (result.status == "succeeded") {
                updateStatus("File closed.");
            }
            else {
                updateStatus("File couldn't be closed.");
            }
        });
    }

    // Get a slice from the file and then call sendSlice.function 
    function getSlice(state) {
        state.file.getSliceAsync(state.counter, function (result) {
            if (result.status == Office.AsyncResultStatus.Succeeded) {
                sendSlice(result.value, state);
            }
            else {

            }
        });
    }

    // Upload the docx file to server after obtaining all the bits from host.
    function onGetAllSlicesSucceeded(docxData) {
        var _temp = myEncodeBase64(docxData);
        //var selId = document.getElementById("IdDocumento");

        var url = 'http://apicoredoc4us.doc4us.com/api/TBR_PlantillaEmpresa/' + selId.value;
        var obj = JSON.stringify({
            "documento": _temp
        });

        $.ajax({
            type: 'PUT',
            data: obj,
            url: url,
            contentType: 'json',
            success: function (data) {
                var _data = data;
                showNotification('Guardado', 'La plantilla se guardo exitosamente!');
            },
            error: function (data) {
                showNotification('error dservivio', data.responseText);
                errorHandler('ERR->AjaxGetAllSlices->' + data.responseText);
            },
            beforeSend: function (xhr) {
                xhr.setRequestHeader('Authorization', 'Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJ1bmlxdWVfbmFtZSI6InBydWViYSIsIm5iZiI6MTU5MTk5MzY3NCwiZXhwIjoxNTkxOTkzOTc0LCJpYXQiOjE1OTE5OTM2NzR9.p9UKbzl-JvAXCVG8zGO_mWI8Pw8QZSfqO6ScrsMorts');
                xhr.setRequestHeader('Content-Type', 'application/json;odata.metadata=minimal;odata.streaming=true');
                xhr.setRequestHeader('accept', '*/*');
            }
        });







        //$.ajax({
        //    type: "POST",
        //    url: "Handler.ashx",
        //    data: myEncodeBase64(docxData),
        //    contentType: "application/json; charset=utf-8",
        //}).done(function (data) {
        //    document.getElementById("documentXmlContent").textContent = data;
        //}).fail(function (jqXHR, textStatus) {
        //});
    }

    function insertText(label) {
        var labelr = String(label);
        Word.run(function (context) {

            // Create a proxy object for the document.
            var thisDocument = context.document;

            // Queue a command to get the current selection.
            // Create a proxy range object for the selection.
            var range = thisDocument.getSelection();
            // Queue a command to replace the selected text.
            range.insertText(labelr + '\n', Word.InsertLocation.replace);

            // Synchronize the document state by executing the queued commands,
            // and return a promise to indicate task completion.
            return context.sync().then(function () {
                console.log('Added a quote from Ralph Waldo Emerson.');
            });
        })
            .catch(function (error) {
                console.log('Error5: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
    }

    //function guardar() {
    //    $(document).ready(function () {
    //        insertText("Es una prueba");
    //        //var api_url = 'https://api.linkpreview.net'
    //        //var key = '5b578yg9yvi8sogirbvegoiufg9v9g579gviuiub8' // not real

    //        //$(".content a").each(function (index, element) {

    //        //    $.ajax({
    //        //        url: api_url + "?key=" + key + " &q=" + $(this).text(),
    //        //        contentType: "application/json",
    //        //        dataType: 'json',
    //        //        success: function (result) {
    //        //            console.log(result);
    //        //        }
    //        //    })
    //        //});
    //    })
    //}





    // Function that writes to a div with id='message' on the page.
    function write(message) {
        document.getElementById('message').innerText += message;
    }



    function loadSampleData() {
        // Run a batch operation against the Word object model.
        Word.run(function (context) {
            // Create a proxy object for the document body.
            var body = context.document.body;

            // Queue a commmand to clear the contents of the body.
            body.clear();
            // Queue a command to insert text into the end of the Word document body.
            body.insertText(
                "Este es un ejemplo de texto para ser trabajado sobre el documento",
                Word.InsertLocation.end);

            // Synchronize the document state by executing the queued commands, and return a promise to indicate task completion.
            return context.sync();
        })
            .catch(errorHandler);
    }

    function hightlightLongestWord() {
        Word.run(function (context) {
            // Queue a command to get the current selection and then
            // create a proxy range object with the results.
            var range = context.document.getSelection();

            // This variable will keep the search results for the longest word.
            var searchResults;

            // Queue a command to load the range selection result.
            context.load(range, 'text');

            // Synchronize the document state by executing the queued commands
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    // Get the longest word from the selection.
                    var words = range.text.split(/\s+/);
                    var longestWord = words.reduce(function (word1, word2) { return word1.length > word2.length ? word1 : word2; });

                    // Queue a search command.
                    searchResults = range.search(longestWord, { matchCase: true, matchWholeWord: true });

                    // Queue a commmand to load the font property of the results.
                    context.load(searchResults, 'font');
                })
                .then(context.sync)
                .then(function () {
                    // Queue a command to highlight the search results.
                    searchResults.items[0].font.highlightColor = '#FFFF00'; // Yellow
                    searchResults.items[0].font.bold = true;
                })
                .then(context.sync);
        })
            .catch(errorHandler);
    }


    function displaySelectedText() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showNotification('The selected text is:', '"' + result.value + '"');
                } else {
                    showNotification('Error6:', result.error.message);
                }
            });
    }

    //$$(Helper function for treating errors, $loc_script_taskpane_home_js_comment34$)$$
    function errorHandler(error) {
        // $$(Always be sure to catch any accumulated errors that bubble up from the Word.run execution., $loc_script_taskpane_home_js_comment35$)$$
        showNotification("Error Handler:", error);
        console.log("Error Handler: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }








    function GestionarPlantilla() {

        // Deshabilitar el guardar documento, ver documento y 
        // prepara las opciones para gestioonar plantillas para tipos de documentos Doc4us-1

        document.getElementById('modulocombinaciondocumento').style.visibility = 'hidden';
        document.getElementById('guardar').style.visibility = 'visible';
        document.getElementById('VerDocumento').style.visibility = 'hidden';
        document.getElementById('VerPreliminar').style.visibility = 'hidden';
        document.getElementById('GuardarDocumento').style.visibility = 'hidden';
        showNotification('Gestión de Plantillas', 'A continuación puede gestionar las plantillas de sus documentos');


    }

    function GestionarDocumento() {
        // deshabilita guardar plantilla y ver documento
        // habilita ver preliminar
        // prepara las opciones para gestioonar plantillas para tipos de documentos Doc4us-2
        document.getElementById('modulocombinaciondocumento').style.visibility = 'visible';
        document.getElementById('guardar').style.visibility = 'hidden';
        document.getElementById('VerDocumento').style.visibility = 'hidden';
        document.getElementById('VerPreliminar').style.visibility = 'visible';
        document.getElementById('GuardarDocumento').style.visibility = 'visible';
        showNotification('Gestión de Documentos', 'A continuación puede gestionar los documentos para radicar');

    }



    function VerPreliminar() {
        showNotification('Ver preliminar', '..........');
        // deshabilita guardar documento
        // habilita ver documento
        // guarda documento
        // combikina correspondencia y muesta documetno casi final
        document.getElementById('modulocombinaciondocumento').style.visibility = 'visible';
        document.getElementById('guardar').style.visibility = 'hidden';
        document.getElementById('VerDocumento').style.visibility = 'visible';
        document.getElementById('GuardarDocumento').style.visibility = 'hidden';
        document.getElementById('VerPreliminar').style.visibility = 'hidden';

    }

    function VerDocumento() {
        showNotification('Ver documento', '..........');
        // recupera documento guardado previamente
        // habilita  ver preliminar
        // habilita guardar documento
        // se oculta asi mismo
        document.getElementById('modulocombinaciondocumento').style.visibility = 'visible';
        document.getElementById('guardar').style.visibility = 'hidden';
        document.getElementById('VerDocumento').style.visibility = 'hidden';
        document.getElementById('VerPreliminar').style.visibility = 'visible';
        document.getElementById('GuardarDocumento').style.visibility = 'visible';
    }

    function GuardarDocumento() {
        //showNotification('guardar documento', '..........');

        // gurda documento
        // habilita ver preliminar
        // oculta ver documento

        getArchivoUrl();
    }

    function getArchivoUrl() {
        //Get the URL of the current file.
        Office.context.document.getFilePropertiesAsync(function (asyncResult) {
            var fileUrl = asyncResult.value.url;
            if (fileUrl == "") {
                showNotification('Archivo', 'archivo no guardado, debe guardar documento ');
            }
            else {
                showNotification('Archivo', fileUrl);
            }
            return fileUrl;
        });
    }


    function insertarEtiqueta() {
        //showNotification('insertarEtiqueta', '..........');
        var selObj = document.getElementById("ListaEtiquetas");

        if (selObj.value) {
            var selValue = selObj.options[selObj.selectedIndex].text;
            writeContent(selValue);
            //insertLabel(selValue);
        }
    }


    function combineText() {
        var selObj = document.getElementById("txtreplace");
        var selValue = selObj.value;
        combineContent(selValue);
    }

    function insertLabel(label) {
        var labelr = String(label);
        Word.run(function (context) {
            debugger;
            // Create a proxy object for the document.
            var thisDocument = context.document;

            // Queue a command to get the current selection.
            // Create a proxy range object for the selection.
            var range = thisDocument.getSelection();
            // Queue a command to replace the selected text.
            range.insertText(labelr + '\n', Word.InsertLocation.replace);

            // Synchronize the document state by executing the queued commands,
            // and return a promise to indicate task completion.
            return context.sync().then(function () {
                console.log('Added a quote from Ralph Waldo Emerson.');
            });
        })
            .catch(function () {
                console.log('Error1: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
    }


    function writeContent(label) {
        var FileXml = "../XMLBase.xml";
        var myOOXMLRequest = new XMLHttpRequest();
        var myTEXT;
        myOOXMLRequest.open('GET', FileXml, false);
        myOOXMLRequest.send();
        if (myOOXMLRequest.status === 200) {
            myTEXT = myOOXMLRequest.responseText;
            myTEXT = myTEXT.replace("TextoAReemplazar", "«" + String(label).trim() + "»");
        }
        Office.context.document.setSelectedDataAsync(myTEXT, { coercionType: 'ooxml' });
    }

    function getFileUrl() {
        Word.run(function (context) {
            var selValue = Office.context.document.url;
            var _App = selValue.split("-");
            var selApp = document.getElementById("App");
            var selId = document.getElementById("IdDocumento");
            selApp.value = _App[1].toString();
            var _SelId = _App[2].split(".");
            selId.value = _SelId[0].toString();
            var thisDocument = context.document;
            var range = thisDocument.getSelection();
            range.insertText(selValue);
            return context.sync().then(function () {
                console.log('Inserto ok');
            });
        });
    }


    function combineContent(label) {
        var labelr = String(label);

        Word.run(function (context) {
            debugger;
            // Create a proxy object for the document.
            debugger;
            var thisDocument = context.document;
            let evenContentControls = thisDocument.contentControls.getById(112233);
            let cssContentControls = thisDocument.contentControls;
            cssContentControls.load();
            evenContentControls.load("length");
            return context.sync().then(function () {
                debugger;
                if (cssContentControls.items.length = 1) {
                    console.log(cssContentControls.items[0].text);
                }
                for (let i = 0; i < evenContentControls.items.length; i++) {
                    // Change a few properties and append a paragraph
                    evenContentControls.items[i].set({
                        color: "red",
                        title: "Odd ContentControl #" + (i + 1),
                        appearance: "Tags"
                    });
                    evenContentControls.items[i].insertParagraph("This is an odd content control", "End");
                }
                return context.sync();
            });
        })
            .catch(function (error) {
                console.log('Error2: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
    }

    function save() {
        var newURL = window.location.protocol + "//" + window.location.host + "/" + window.location.pathname + window.location.search
        console.log('Url info: ' + newURL);
    }

    //function combineContent(label) {
    //var labelr = String(label);
    //// Adds title and colors to odd and even content controls and changes their appearance.
    //    debugger;
    //Word.run(function (context) {
    //// Gets the complete sentence (as range) associated with the insertion point.
    //let evenContentControls = context.document.contentControls.getByTag("even");
    //let oddContentControls = context.document.contentControls.getByTag("odd");
    //evenContentControls.load("length");
    //oddContentControls.load("length");

    //await context.sync();

    //for (let i = 0; i < evenContentControls.items.length; i++) {
    //    // Change a few properties and append a paragraph
    //    evenContentControls.items[i].set({
    //    color: "red",
    //    title: "Odd ContentControl #" + (i + 1),
    //    appearance: "Tags"
    //    });
    //    evenContentControls.items[i].insertParagraph("This is an odd content control", "End");
    //    }
    //  });
    //}

    function getEtiquetas() {
        showNotification('getEtiquetas', '..........');

        var _Etiqueta = $("#ListaEtiquetas");
        var _IdClasEtiq = document.getElementById("TipoEtiqueta").value;
        var url = "http://apicoredoc4us.doc4us.com/api/TS_Etiqueta/GetAll?%24filter=id_TS_ClasificacionEtiqueta%20eq%20" + _IdClasEtiq;
        $.ajax({
            type: 'GET',
            url: url,
            contentType: 'json',
            success: function (data) {
                _Etiqueta.find('option').remove();
                $.each(data, function (key, registro) {
                    _Etiqueta.append('<option value="' + registro.id + '">' + registro.nombre + '</option>');
                });
            },
            error: function (data) {
                _Etiqueta.find('option').remove();
                //var err = data.err;
                showNotification("Error Servicio Lista Etiquetas", data.responseText);
                errorHandler('ERR->GetEtiqueta->' + data.responseText);
            },
            beforeSend: function (xhr) {
                xhr.setRequestHeader('Authorization', 'Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJ1bmlxdWVfbmFtZSI6InBydWViYSIsIm5iZiI6MTU5MTk5MzY3NCwiZXhwIjoxNTkxOTkzOTc0LCJpYXQiOjE1OTE5OTM2NzR9.p9UKbzl-JvAXCVG8zGO_mWI8Pw8QZSfqO6ScrsMorts');
            }
        });
    }

    function getClasificacionEtiquetas() {
        var _TipoEtiqueta = $("#TipoEtiqueta");
        var url = 'http://apicoredoc4us.doc4us.com/api/TS_ClasificacionEtiqueta/GetAll';
        $.ajax({
            type: 'GET',
            url: url,
            contentType: 'json',
            success: function (data) {
                _TipoEtiqueta.find('option').remove();
                _TipoEtiqueta.append('<option value=-1">Seleccione</option>');
                $.each(data, function (key, registro) {
                    _TipoEtiqueta.append('<option value="' + registro.id + '">' + registro.nombre + '</option>');
                });
            },
            error: function (data) {
                showNotification("Error Clasifica Etiquetas", data.responseText);
                //var err = data.err;
                errorHandler('ERR->GetClasifica->' + data.responseText);
            },
            beforeSend: function (xhr) {
                xhr.setRequestHeader('Authorization', 'Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJ1bmlxdWVfbmFtZSI6InBydWViYSIsIm5iZiI6MTU5MTk5MzY3NCwiZXhwIjoxNTkxOTkzOTc0LCJpYXQiOjE1OTE5OTM2NzR9.p9UKbzl-JvAXCVG8zGO_mWI8Pw8QZSfqO6ScrsMorts');
            }
        });
        //Word.run(function (context) {
        //    var selValue = Office.context.document.url;
        //    if (selValue !== "") {
        //        var _App = selValue.split("-");
        //        var selApp = document.getElementById("App");
        //        var selId = document.getElementById("Id");
        //        var _SelId = _App[2].split(".");
        //        selId.value = _SelId[0].toString();
        //        selApp.value = _App[1].toString();
        //        if (selApp.value === "2") {
        //            $("#TI2").css("display", "none");
        //        }
        //        else if (selApp.value === "1") {
        //            $("#TI3").css("display", "none");
        //        }
        //        else {
        //            $("#TI2").css("display", "none");
        //            $("#TI3").css("display", "none");
        //        }
        //    }
        //    else {
        //        $("#TI2").css("display", "none");
        //        $("#TI3").css("display", "none");
        //    }

        //    return context.sync().then(function () {
        //        console.log('Inserto ok');
        //    });
        //});
    }



})();
