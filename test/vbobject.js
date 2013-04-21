if(! vb) {
    // load `vb.coffee` module when running under Node
    require("coffee-script");
    var vb = require("../vb");
}

// == SECTION VBObject basic

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
print(o.get('Bla'));
// => bla


// == SECTION VBObject call

var ns = new vb.VBObject({type: 'Namespace',
                          attrs: {
                            DoCmd: new vb.VBObject({
                              type: 'DoCmd',
                              attrs: {
                                OpenQuery: function (a, b) {
                                  print ('OpenQuery(' + a + ', ' + b + ')');
                                  return a;
                                }
                              }})}});

print(ns.get('DoCmd').get('OpenQuery')('AAA', 'BBB'));
// => OpenQuery(AAA, BBB)
// => AAA

print(ns.get('DoCmd').dot('OpenQuery')('AAA', 'BBB'));
// => OpenQuery(AAA, BBB)
// => AAA
