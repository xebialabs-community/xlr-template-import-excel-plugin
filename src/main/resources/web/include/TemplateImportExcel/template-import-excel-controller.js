/*
 * Copyright 2019 XEBIALABS
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
 */
(function(module) {
    'use strict';

    function TemplateImportExcelController($scope, $httpBackend) {

        var vm = this;

        vm.importExcel = importExcel;

        vm.httpBackend = $httpBackend;

        function callback() {
            alert("Executing callback()");
        }

        function toBase64(indata) {
        //  Need a Base64 encoder here.  bota conversion gets corrupted on upload
            return indata;
        }

        function importExcel(excelFile) {
            var reader = new FileReader();
            reader.onload = function(e) {
                var payload = {"data":toBase64(e.target.result)}
                // var payload = new Uint8Array(e.target.result);
                vm.httpBackend('POST', '/api/extension/templateImportExcel/import', JSON.stringify(payload), callback, {'Content-Type':'application/json'});
            }
            reader.readAsText(excelFile.files[0]);
        }
    }

    TemplateImportExcelController.$inject = ['$scope', '$httpBackend'];

    module.controller('templateImportExcelController', TemplateImportExcelController);

} ) (angular.module('xlrelease'));
