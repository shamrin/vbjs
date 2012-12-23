## JavaScript API

```javascript
var fs = require('fs')
var vb = require('vb');

var Close = function() {
    console.log('Close');
};
var OpenForm = function(name) {
    console.log('OpenForm ' + name);
}

var m = vb.runModule(fs.readFileSync('nwind_Startup.bas'),
                     { dotobj:
                         { DoCmd:
                            { dotobj:
                               { Close: Close, OpenForm: OpenForm }}}});

m.CloseForm();
// -> Close
// -> OpenForm 'Main Switchboard'
```

`nwind_Startup.bas` was taken from Microsoft Access&reg; Northwind sample database `Startup` module.

## Command line usage

```sh
$ cat nwind_Startup.bas | coffee vb.coffee
```

Result:

```javascript
return {
    ..
    'CloseForm': function () {
        ns('DoCmd').dot('Close')();
        ns('DoCmd').dot('OpenForm')('Main Switchboard');
    },
    ...
};
```
