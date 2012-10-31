{evaluate, compile, VBRuntimeError} = require '../vb'
assert = require 'assert'

run = (expr, me={}, us={}, fns={}) ->
    namespace =
        _Us: dot: (name) ->
                    unless us[name]?
                      throw new VBRuntimeError "'#{name}' not found in a _Us"
                    us[name]
        Me: dot: (name) ->
                    unless me[name]?
                      throw new VBRuntimeError "'#{name}' not found in a Me"
                    me[name]
    for k, v of fns
        namespace[k] = v
    evaluate expr, namespace

eq = (expected, actual, msg) -> assert.strictEqual actual, expected, msg

nancy = FirstName: 'Nancy', LastName: 'Davolio'
fns =
    Abs: (ns, expr) -> Math.abs(expr)
    Sum: (ns, expr) ->
        # TODO implement Sum in VBA, using CurrentDb.OpenRecordset or DBEngine
        field = {'[Field]': 'Field'}[expr]
        sum = 0
        for val in ns('_Us').dot(field) then sum += val
        sum

suite 'Expressions -', ->
    setup ->
        # nothing here
    test 'basic', ->
        eq 'Nancy, Davolio', run '[FirstName]&", "&[LastName]', nancy
    test 'basic whitespace', ->
        eq 'Nancy Davolio', run '[FirstName] & " " & [LastName]', nancy
    test 'identifier with whitespace', ->
        eq 'Nancy', run '[First name]', 'First name': 'Nancy'
    test 'identifier operators', ->
        eq 'Total: 10', run '"Total: " & [Some Subform].[Form]![Subtotal]',
                        'Some Subform':
                            dot: (f) -> {Form:
                                            bang: (g) -> {Subtotal: 10}[g]}[f]
    test 'addition', ->
        eq 20, run '[Subtotal]+[Freight]', Subtotal: 13, Freight: 7
    test 'functions', ->
        eq 30, run 'Abs([Field])', {Field: -30}, {}, fns
    test 'lazy functions', ->
        eq 40, run 'Sum([Field])', {}, {Field: [10, 10, 20]}, fns
    test 'arithmetic', ->
        eq 5, run '[A] / [B] + [C] * [D] - [E]', A: 9, B: 3, C: 2, D: 3, E: 4
    test 'float', ->
        eq 2.0, run '[A] + 0.5', A: 1.5
    test 'nested calls', ->
        eq 60, run 'Abs(Sum([Field]))', {}, {Field: [10, 20, -90]}, fns
    test 'unknown function error' , ->
        assert.throws (-> run 'Foo([F])', {F: 123}), VBRuntimeError
    test 'unknown me.field error', ->
        assert.throws (-> run '[Field]'), VBRuntimeError
    test 'unknown us.field error', ->
        assert.throws (-> run 'Sum([Field])', {}, {}, fns), VBRuntimeError
    test 'generated code', ->
        eq "function anonymous(ns) {\nvar me = ns('Me').dot; return me('Field');\n}",
           compile('[Field]').toString()
