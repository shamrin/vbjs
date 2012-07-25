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

{strictEqual: eq} = require 'assert'
{evaluate} = require '../vbjs'

ev = (expr) ->
    evaluate expr, {FirstName: 'Nancy', LastName: 'Davolio'}

suite 'Expressions', ->
    setup ->
        # nothing here
    test 'basic', ->
        eq 'Nancy Butley', ev '[FirstName]&" "&[LastName]'
