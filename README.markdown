## Current status

Compiling `Startup` module from Microsoft Access&reg; Northwind sample database:

```sh
$ cat nwind_Startup.bas | coffee vb.coffee
```

Result:

```javascript
return {
    'OpenStartup': function () {
        if (ns('IsItAReplica')()) {
            ns('DoCmd').dot('Close')();
        } else {
            if (ns('CurrentDb')().dot('Properties')('StartupForm') === 'Startup' || ns('CurrentDb')().dot('Properties')('StartupForm') === 'Form.Startup') {
                ns('Forms').bang('Startup').bang('HideStartupForm').let(false);
            } else {
                ns('Forms').bang('Startup').bang('HideStartupForm').let(true);
            }
        }
        return;
        if (ns('Err') === ns('conPropertyNotFound')) {
            ns('Forms').bang('Startup').bang('HideStartupForm').let(true);
        }
    },
    'HideStartupForm': function () {
        if (ns('Forms').bang('Startup').bang('HideStartupForm')) {
            ns('CurrentDb')().dot('Properties')('StartupForm').let('Main SwitchBoard');
        } else {
            ns('CurrentDb')().dot('Properties')('StartupForm').let('Startup');
        }
        return;
        if (ns('Err') === ns('conPropertyNotFound')) {
            ns('db').dot('Properties').dot('Append')(ns('prop'));
        }
    },
    'CloseForm': function () {
        ns('DoCmd').dot('Close')();
        ns('DoCmd').dot('OpenForm')('Main Switchboard');
    },
    'IsItAReplica': function () {
        ns('blnReturnValue').let(false);
        if (ns('CurrentDb')().dot('Properties')('Replicable') === 'T') {
            ns('blnReturnValue').let(true);
        } else {
            ns('blnReturnValue').let(false);
        }
        ns('IsItAReplica').let(ns('blnReturnValue'));
        return;
    }
};
```
