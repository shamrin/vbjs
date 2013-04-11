// load CoffeeScript module to teach `require()` about `.coffee` modules
try { require("coffee-script"); } catch (e) { }

var vb = require("../vb");

var o = new vb.VBObject({type: 'TextBox',
                         attrs: {visible: true, value: 'foo'},
                         default: 'value'});

print(o.dot('visible').get()); // => true
print(o.get('Visible')); // => true
print(o.dot('value').get()); // => foo
print(o.get('value')); // => foo
print(o.get()); // => foo

o.let('value', 'bar');
print(o.get()); // => bar

o.dot('Value').let('baz');
print(o.get()); // => baz

o.dot('Bla').get();
// => Error: VBRuntimeError: TextBox has no attribute 'Bla'

o.get('Bla');
// => Error: VBRuntimeError: TextBox has no attribute 'Bla'

o.dot('Bla').let('bla');
// => Error: VBRuntimeError: TextBox has no attribute 'Bla'

