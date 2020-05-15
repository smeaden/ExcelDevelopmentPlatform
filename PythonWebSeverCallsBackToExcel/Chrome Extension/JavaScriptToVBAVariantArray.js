'use strict';

// module exporting for node.js and browsers with thanks to
// https://www.matteoagosti.com/blog/2013/02/24/writing-javascript-modules-for-both-browser-and-node/

(function () {
    var JavaScriptToVBAVariantArray = (function () {
        var JavaScriptToVBAVariantArray = function (options) {
            var pass; //...
        };


        JavaScriptToVBAVariantArray.prototype.testPersistVar = function testPersistVar() {
            try {
                var payload;
                //payload = "Hello World";
                //payload = false;
                //payload = 655.35;
                payload = new Date(1989, 9, 16, 12, 0, 0);
                var payloadEncoded = persistVar(payload);
                return payloadEncoded;
            }
            catch (err) {
                console.log(err.message);
            }
        };

        JavaScriptToVBAVariantArray.prototype.testPersistGrid = function testPersistGrid() {
            try {
                var rows = 2;
                var columns = 4;
                var arr = this.createGrid(rows, columns);
                arr[0][0] = "Hello World";
                arr[0][1] = true;
                arr[0][2] = false;
                arr[0][3] = null;

                arr[1][0] = 65535;
                arr[1][1] = 7.5;
                arr[1][2] = new Date(1989, 9, 16, 12, 0, 0);
                arr[1][3] = new Error(2042);

                var payloadEncoded = this.persistGrid(arr, rows, columns);
                return payloadEncoded;
            }
            catch (err) {
                console.log(err.message);
            }
        };


		JavaScriptToVBAVariantArray.prototype.persistGrid = function persistGrid(grid, rows, columns) {
			try {
				/* Opening sequence of bytes is a reduced form of SAFEARRAY and SAFEARRAYBOUND
				 * SAFEARRAY       https://docs.microsoft.com/en-gb/windows/win32/api/oaidl/ns-oaidl-safearray
				 * SAFEARRAYBOUND  https://docs.microsoft.com/en-gb/windows/win32/api/oaidl/ns-oaidl-safearraybound
				 */

				var payloadEncoded = new Uint8Array(20);

				// vbArray + vbVariant, lo byte, hi byte
				payloadEncoded[0] = 12; payloadEncoded[1] = 32;

				// number of dimensions, lo byte, hi byte
				payloadEncoded[2] = 2; payloadEncoded[3] = 0;

				// number of columns, 4 bytes, least significant byte first
				payloadEncoded[4] = columns % 256; payloadEncoded[5] = Math.floor(columns / 256);
				payloadEncoded[6] = 0; payloadEncoded[7] = 0;

				// columns lower bound (safearray)
				payloadEncoded[8] = 1; payloadEncoded[9] = 0;
				payloadEncoded[10] = 0; payloadEncoded[11] = 0;

				// number of rows, 4 bytes, least significant byte first
				payloadEncoded[12] = rows % 256; payloadEncoded[13] = Math.floor(rows / 256);
				payloadEncoded[14] = 0; payloadEncoded[15] = 0;

				// rows lower bound (safearray)
				payloadEncoded[16] = 1; payloadEncoded[17] = 0;
				payloadEncoded[18] = 0; payloadEncoded[19] = 0;

				var elementBytes;
				for (var colIdx = 0; colIdx < columns; colIdx++) {
					for (var rowIdx = 0; rowIdx < rows; rowIdx++) {
						elementBytes = this.persistVar(grid[rowIdx][colIdx]);
						var arr = [payloadEncoded, elementBytes];

						payloadEncoded = this.concatArrays(arr); // Browser
					}
				}
				return payloadEncoded;
			}
			catch (err) {
				console.log(err.message);
			}
		};



        JavaScriptToVBAVariantArray.prototype.concatArrays = function concatArrays(arrays) {
            // With thanks to https://javascript.info/arraybuffer-binary-arrays


            // sum of individual array lengths
            let totalLength = arrays.reduce((acc, value) => acc + value.length, 0);

            if (!arrays.length) return null;

            let result = new Uint8Array(totalLength);

            // for each array - copy it over result
            // next array is copied right after the previous one
            let length = 0;
            for (let array of arrays) {
                result.set(array, length);
                length += array.length;
            }

            return result;
        };

        JavaScriptToVBAVariantArray.prototype.createGrid = function createGrid(rows, columns) {
            try {
                return Array.from(Array(rows), () => new Array(columns));
            }
            catch (err) {
                console.log(err.message);
            }
        };

        JavaScriptToVBAVariantArray.prototype.persistVar = function persistVar(v) {
            try {

                if (v === null) {
                    // return a Null
                    var nullVt = new Uint8Array(2);
                    nullVt[0] = 1;
                    return nullVt;

                } else if (v instanceof Error) {

                    return this.persistError(v);

                } else if (typeof v === 'undefined') {
                    return new Uint8Array(2); // return an Empty

                } else if (typeof v === "boolean") {
                    // variable is a boolean
                    return this.persistBool(v);
                } else if (typeof v.getMonth === "function") {
                    // variable is a Date
                    return this.persistDate(v);
                } else if (typeof v === "string") {
                    // variable is a boolean
                    return this.persistString(v);
                } else if (typeof v === "number") {
                    // variable is a number
                    return this.persistNumber(v);
                }

            }
            catch (err) {
                console.log(err.message);
            }
        };

        JavaScriptToVBAVariantArray.prototype.persistError = function persistError(v) {
            try {
                var errorVt = new Uint8Array(6); // return a vtError
                errorVt[0] = 10; errorVt[4] = 10; errorVt[5] = 128;

                var errorNumber;
                try {
                    errorNumber = parseInt(v.message);
                }
                catch (err) {
                    errorNumber = 2000;
                    console.log(err.message);
                }
                errorVt[2] = errorNumber % 256; errorVt[3] = Math.floor(errorNumber / 256);

                return errorVt;
            }
            catch (err) {
                console.log(err.message);
            }
        };

        JavaScriptToVBAVariantArray.prototype.persistNumber = function persistNumber(v) {
            try {
                var bytes;
                if (Number.isInteger(v)) {
                    bytes = new Uint8Array(6);
                    bytes[0] = 3; bytes[1] = 0;  // VarType 5 = Long
                    bytes[2] = v % 256; v = Math.floor(v / 256);
                    bytes[3] = v % 256; v = Math.floor(v / 256);
                    bytes[4] = v % 256; v = Math.floor(v / 256);
                    bytes[5] = v % 256;

                } else {
                    bytes = this.persistDouble(v, 5);
                }
                return bytes;
            }
            catch (err) {
                console.log(err.message);
            }
        };

        JavaScriptToVBAVariantArray.prototype.persistDate = function persistDate(v) {
            try {
                // convert JavaScript 1970 base to VBA 1900 base
                // https://stackoverflow.com/questions/46200980/excel-convert-javascript-unix-timestamp-to-date/54153878#answer-54153878
                var xlDate = v / (1000 * 60 * 60 * 24) + 25569;
                return this.persistDouble(xlDate, 7);
            }
            catch (err) {
                console.log(err.message);
            }
        };

        JavaScriptToVBAVariantArray.prototype.persistDouble = function persistDouble(v, vt) {
            try {
                var bytes;
                bytes = new Uint8Array(10);
                bytes[0] = vt; bytes[1] = 0;  // VarType 5 = Double or 7 = Date
                var doubleAsBytes = this.doubleToByteArray(v);
                for (var idx = 0; idx < 8; idx++) {
                    bytes[2 + idx] = doubleAsBytes[idx];
                }
                return bytes;
            }
            catch (err) {
                console.log(err.message);
            }
        };

        JavaScriptToVBAVariantArray.prototype.doubleToByteArray = function doubleToByteArray(number) {
            try {
                // https://stackoverflow.com/questions/25942516/double-to-byte-array-conversion-in-javascript/25943197#answer-39515587
                var buffer = new ArrayBuffer(8);         // JS numbers are 8 bytes long, or 64 bits
                var longNum = new Float64Array(buffer);  // so equivalent to Float64

                longNum[0] = number;

                return Array.from(new Int8Array(buffer));
            }
            catch (err) {
                console.log(err.message);
            }
        };

        JavaScriptToVBAVariantArray.prototype.persistString = function persistString(v) {
            try {
                var strlen = v.length;
                var bytes = new Uint8Array(strlen + 4);
                bytes[0] = 8; bytes[1] = 0;  // VarType 8 = String
                bytes[2] = strlen % 256; bytes[3] = Math.floor(strlen / 256);
                for (var idx = 0; idx < strlen; idx++) {
                    bytes[idx + 4] = v.charCodeAt(idx);
                }
                return bytes;
            }
            catch (err) {
                console.log(err.message);
            }
        };

        JavaScriptToVBAVariantArray.prototype.persistBool = function persistBool(v) {
            try {
                var bytes = new Uint8Array(4);
                bytes[0] = 11; bytes[1] = 0;   // VarType 11 = Boolean
                if (v === true) {
                    bytes[2] = 255; bytes[3] = 255;
                } else {
                    bytes[2] = 0; bytes[3] = 0;
                }
                return bytes;
            }
            catch (err) {
                console.log(err.message);
            }
        };

        return JavaScriptToVBAVariantArray;
    })();

    if (typeof module !== 'undefined' && typeof module.exports !== 'undefined')
        module.exports = JavaScriptToVBAVariantArray;
    else
        window.JavaScriptToVBAVariantArray = JavaScriptToVBAVariantArray;
})();