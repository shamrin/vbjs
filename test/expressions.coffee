# Other tests to add (from nwind.mdb):
#    '[Quarterly Orders Subform]![Total]'
#    '"Grand Total for " & [Forms]![Quarterly Orders]![Quarterly Orders Subform].[Form]![Year]'
#    'NZ(Sum([Qtr 1]))'
#    '[TotalQ1]+[TotalQ2]+[TotalQ3]+[TotalQ4]'

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
        eq 10, run '[Some Subform].[Form]![Subtotal]',
                   'Some Subform':
                       Form: __default: (v) -> {Subtotal: 10}[v]
    test 'addition', ->
        eq 20, run '[Subtotal]+[Freight]', Subtotal: 13, Freight: 7
    test 'functions', ->
        eq 30, run 'Sum([ExtendedPrice])', {}, {ExtendedPrice: [4, 10, 16]},
                   Sum: (Me, Us, expr) ->
                       a = Us[{'[ExtendedPrice]': 'ExtendedPrice'}[expr]]
                       s = 0
                       for e in a then s += e
                       s
