function getData() {
    var _TipoEtiqueta = $("#TipoEtiqueta");
    var url = 'http://apicoreibpmns.doc4us.com/api/TS_ClasificacionEtiqueta/GetAll';
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
            var err = data.err;
        },
        beforeSend: function (xhr) {
            xhr.setRequestHeader('Authorization', 'Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJ1bmlxdWVfbmFtZSI6InBydWViYSIsIm5iZiI6MTU5MTk5MzY3NCwiZXhwIjoxNTkxOTkzOTc0LCJpYXQiOjE1OTE5OTM2NzR9.p9UKbzl-JvAXCVG8zGO_mWI8Pw8QZSfqO6ScrsMorts');
        }
    });
    Word.run(function (context) {
        var selValue = Office.context.document.url;
        if (selValue !== "") {
            var _App = selValue.split("-");
            var selApp = document.getElementById("App");
            var selId = document.getElementById("Id");
            var _SelId = _App[2].split(".");
            selId.value = _SelId[0].toString();
            selApp.value = _App[1].toString();
            if (selApp.value === "2") {
                $("#TI2").css("display", "none");
            }
            else if (selApp.value === "1") {
                $("#TI3").css("display", "none");
            }
            else {
                $("#TI2").css("display", "none");
                $("#TI3").css("display", "none");
            }
        }
        else {
            $("#TI2").css("display", "none");
            $("#TI3").css("display", "none");
        }

        return context.sync().then(function () {
            console.log('Inserto ok');
        });
    });
};

function getDisponibles() {
    var _Etiqueta = $("#ListaCampos");
    var _IdClasEtiq = document.getElementById("TipoEtiqueta").value;
    var url = "http://apicoreibpmns.doc4us.com/api/TS_Etiqueta/GetAll?%24filter=id_TS_ClasificacionEtiqueta%20eq%20" + _IdClasEtiq;
    $.ajax({
        type: 'GET',
        url: url,
        contentType: 'json',
        success: function (data) {
            _Etiqueta.find('option').remove();
            $.each(data, function (key, registro) {
                _Etiqueta.append('<option value="' + registro.id + '">' + registro.etiqueta + '</option>');
            });
        },
        error: function (data) {
            _Etiqueta.find('option').remove();
            var err = data.err;
        },
        beforeSend: function (xhr) {
            xhr.setRequestHeader('Authorization', 'Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJ1bmlxdWVfbmFtZSI6InBydWViYSIsIm5iZiI6MTU5MTk5MzY3NCwiZXhwIjoxNTkxOTkzOTc0LCJpYXQiOjE1OTE5OTM2NzR9.p9UKbzl-JvAXCVG8zGO_mWI8Pw8QZSfqO6ScrsMorts');
        }
    });
};

(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {

            //document.getElementById("getFileAsync").click(getFileAsyncInternal);

            if (Office.context.requirements.isSetSupported('WordApi', '1.1')) {
                // Do something that is only available via the new APIs
                //$('#guardar').click(guardar());
                $('#checkhov').click(insertChekhovQuoteAtTheBeginning);
                $('#proverb').click(insertChineseProverbAtTheEnd);
                $('#supportedVersion').html('Esta Add-ins de Doc4us debe ser usado para Word 2016 o Superior.');
                //$('#getFileAsync').click(guardar);
                $('#guardar').click(guardar);




            }
            else {
                // Just letting you know that this code will not work with your version of Word.
                $('#supportedVersion').html('This code requires Word 2016 or later.');
            }


            var DropdownHTMLElements = document.querySelectorAll('.ms-Dropdown');
            for (var i = 0; i < DropdownHTMLElements.length; ++i) {
                var Dropdown = new fabric['Dropdown'](DropdownHTMLElements[i]);
            }
            var PivotElements = document.querySelectorAll(".ms-Pivot");
            for (var j = 0; j < PivotElements.length; j++) {
                new fabric['Pivot'](PivotElements[j]);
            }
            var CommandButtonElements = document.querySelectorAll(".ms-CommandButton");
            for (var k = 0; k < CommandButtonElements.length; k++) {
                new fabric['CommandButton'](CommandButtonElements[k]);
            }
            var TextFieldElements = document.querySelectorAll(".ms-TextField");
            for (var l = 0; l < TextFieldElements.length; l++) {
                new fabric['TextField'](TextFieldElements[l]);
            }


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

    function guardar() {
        var _UrlFile;
        Word.run(function (context) {
            var thisDocument = context.document;
            _UrlFile = Office.context.document.url;
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
                console.log("Error: " + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
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
        var url = 'http://apicoreibpmns.doc4us.com/api/TP_DetalleTipoDato';
        var selId = document.getElementById("Id");
        var obj = JSON.stringify({
            "id": selId,
            "id_TS_TipoDato": 1,
            "sid_TS_TipoDato": "plantilla generica",
            "codigo": "1521",
            "nombre": "plantilla generica",
            "detalle": "plantilla generica",
            "detallePlantilla": _temp,
            "estado": true,
            "filtro": "",
            "_ippublica": "",
            "_nombremaquina": "",
            "_usuario": "",
            "_ipdetrasproxy": "",
            "_browser": "",
            "_accion": "",
            "_sessionid": ""
        });



        $.ajax({
            type: 'PUT',
            data: obj,
            url: url,
            contentType: 'json',
            success: function (data) {
                var _data = data;
            },
            error: function (data) {
                var err = data.err;
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
                console.log('Error: ' + JSON.stringify(error));
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

    function insertChekhovQuoteAtTheBeginning() {
        Word.run(function (context) {

            // Create a proxy object for the document body.
            var body = context.document.body;

            // Queue a command to insert text at the start of the document body.
            body.insertText('"Knowledge is of no value unless you put it into practice."\n', Word.InsertLocation.start);

            // Synchronize the document state by executing the queued commands,
            // and return a promise to indicate task completion.
            return context.sync().then(function () {
                console.log('Added a quote from Anton Chekhov.');
            });
        })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
    }




    function insertChineseProverbAtTheEnd() {
        Word.run(function (context) {

            // Create a proxy object for the document body.
            var body = context.document.body;

            // Queue a command to insert text at the end of the document body.
            body.insertText('"To know the road ahead, ask those coming back."\n', Word.InsertLocation.end);

            // Synchronize the document state by executing the queued commands,
            // and return a promise to indicate task completion.
            return context.sync().then(function () {
                console.log('Added a quote from a Chinese proverb.');
            });
        })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
    }

    // Function that writes to a div with id='message' on the page.
    function write(message) {
        document.getElementById('message').innerText += message;
    }



})();
