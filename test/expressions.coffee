#inputs = [
#    '[FirstName]&" "&[LastName]'
#    '[FirstName] & " " & [LastName]'
#    '[Orders Subform].[Form]![OrderSubtotal]'
#    '[Subtotal]+[Freight]'
#    'Sum([ExtendedPrice])'
#    '[Quarterly Orders Subform]![Total]'
#    '"Grand Total for " & [Forms]![Quarterly Orders]![Quarterly Orders Subform].[Form]![Year]'
#    'NZ(Sum([Qtr 1]))'
#    '[TotalQ1]+[TotalQ2]+[TotalQ3]+[TotalQ4]'
#]

{strictEqual} = require 'assert'
{evaluate} = require '../vbjs'

run = (expr) ->
    evaluate expr, {FirstName: 'Nancy', LastName: 'Davolio'}

eq = (expected, actual, message) -> strictEqual actual, expected, message

suite 'Expressions', ->
    setup ->
        # nothing here
    test 'basic', ->
        eq 'Nancy Butley', run '[FirstName]&" "&[LastName]'
