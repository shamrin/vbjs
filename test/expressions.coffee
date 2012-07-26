# Other tests to add (from nwind.mdb):
#    'NZ(Sum([Qtr 1]))'

{evaluate} = require '../vbjs'
{strictEqual} = require 'assert'

run = (expr, Me, Us, functions) -> evaluate expr, Me, Us, functions
eq = (expected, actual, message) -> strictEqual actual, expected, message

nancy = FirstName: 'Nancy', LastName: 'Davolio'

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
                            Form: __default: (v) -> {Subtotal: 10}[v]
    test 'addition', ->
        eq 20, run '[Subtotal]+[Freight]', Subtotal: 13, Freight: 7
    test 'functions', ->
        eq 30, run 'Abs([Field])', {Field: -30}, {},
                   Abs: (Me, Us, expr) -> Math.abs(expr)
    test 'lazy functions', ->
        eq 40, run 'Sum([ExtendedPrice])', {}, {ExtendedPrice: [10, 10, 20]},
                   Sum: (Me, Us, expr) ->
                       field = {'[ExtendedPrice]': 'ExtendedPrice'}[expr]
                       sum = 0
                       for val in Us[field] then sum += val
                       sum
    test 'long addition', ->
        eq 50, run '[X] + [Y] + [Z]', X: 10, Y: 20, Z: 20
