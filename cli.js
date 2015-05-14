/*
The MIT License (MIT)

Copyright (c) 2015 Michel Petit

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
*/

(function(){
    var argv;



    /**
     * @constructor 
     */
    this.init = function(){
        this.parse();
        this.dispatch();
    };
    


    /**
     * Parse arguments passed to the CLI. 
     */
    this.parse = function(){
        this.argv = require('minimist')(
            process.argv.slice(2), 
            {
                boolean: [
                    "z", 
                    "zero",
                    "j", 
                    "json", 
                    "version", 
                    "help",
                    "h",
                    "stdin"
                ]
            }
        );
    };



    /**
     * Converts passed values
     *
     * If an exception occurs, exit CLI. 
     */
    this.convert = function(input){
        var SpreadsheetColumn = require('./index.js');
        var sc = new SpreadsheetColumn({
            zero: this.argv.z || this.argv.zero
        });

        try {
            var result;

            if(this.argv._.length){
                result = sc.fromAny(this.argv._.join(' '));
            } else if(typeof(input) === 'string') {
                result = sc.fromAny(input);
            }

            if(this.argv.j || this.argv.json){
                console.log(JSON.stringify(result));
            } else {
                for(var i in result){
                    console.log(result[i].original + ' => ' + result[i].converted);
                }
            }
        } catch(err){
            console.error(err);
            process.exit(1);
        }
    };



    /**
     * Display some version information about this project.
     *
     * It uses its own `package.json` to get its data.
     */
    this.version = function(){
        var path = require('path');
        require('fs').readFile(
            __dirname + path.sep + 'package.json', 
            {
                encoding: 'UTF-8',
                flag: 'r'
            },
            function(err, data){
                if(err){
                    console.error(err);
                    process.exit(1);
                } else {
                    var d = JSON.parse(data);
                    var sprintf = require('sprint');
                    if(typeof(d.author) === 'object'){
                        console.log(sprintf(
                            "%s version %s by %s <%s>",
                            d.name, 
                            d.version,
                            d.author.name,
                            d.author.email
                        ));
                    } else {
                        console.log(sprintf(
                            "%s version %s by %s",
                            d.name,
                            d.version,
                            d.author
                        ));
                    }
                    console.log(d.license + ' license');
                }
            }
        );
    };



    /**
     * display help message about CLI script. 
     */
    this.help = function(){
        console.log('Converts spreadsheet column name to number and number to spreadsheet column name.');
        console.log('Synopsis: spreadsheet-column [option] [ARG1 ARG2…]');
        console.log(" -z, --zero       Numbers start from zero in place of one.");
        console.log(" -j, --json       Outputs result as JSON.");
        console.log("     --stdin      Accepts data from standard input (pipe or keyboard).");
        console.log("     --version    Version information.");
        console.log(" -h, --help       Shows help message.");
    };



    /**
     * Converts each data passed using `stdin` (pipe or keyboard). 
     */
    this.acceptFromPipe = function(success, fail){
        var input = '';

        process.stdin.setEncoding('utf8');

        process.stdin.on('readable', function(){
            chunk = process.stdin.read();

            if (chunk !== null && typeof(chunk) !== 'undefined') {
                input += chunk = chunk.trim(/[\n]+/g);
                success(chunk);
            }
        });

        process.stdin.on('end', function(){
            if(input.trim().length){
                success(input);
            } else {
                fail();
            }
        });
    };



    /**
     * Runs the right task for the right arguments passed.
     */
    this.dispatch = function(){
        if(this.argv._.length){
            this.convert();
        } else if(this.argv.version){
            this.version();
        } else if(this.argv.h || this.argv.help){
            this.help();
        } else if(this.argv.stdin){
            this.acceptFromPipe(this.convert, this.help);
        } else {
            this.help();
        }
    };



    /**
     * Run the script. 
     */
    this.init();

})();
