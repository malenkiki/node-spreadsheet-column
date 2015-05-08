#!/usr/bin/env node

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

    this.init = function(){
        this.parse();
        this.dispatch();
    };
    
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
                    "h"
                ]
            }
        );
    };

    this.convert = function(){
        var SpreadsheetColumn = require('./index.js');
        var sc = new SpreadsheetColumn({
            zero: this.argv.z || this.argv.zero
        });

        try {
            var result = sc.fromAny(this.argv._.join(' '));

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
                    console.log(d.name +' version '+ d.version + ' by ' + d.author);
                    console.log(d.license + ' license');
                }
            }
        );
    };

    this.help = function(){
        console.log('Converts spreadsheet column name to number and number to spreadsheet column name.');
        console.log('Synopsis: spreadsheet-column [option] [ARG1 ARG2…]');
        console.log(" -z, --zero       Numbers start from zero in place of one.");
        console.log(" -j, --json       Outputs result as JSON.");
        console.log("     --version    Version information.");
        console.log(" -h, --help       Shows help message.");
    };

    this.dispatch = function(){
        if(this.argv._.length){
            this.convert();
        } else if(this.argv.version){
            this.version();
        } else if(this.argv.h || this.argv.help){
            this.help();
        } else {
            this.help();
        }
    };

    this.init();

})();
