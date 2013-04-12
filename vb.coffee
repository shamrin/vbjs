{isEqual, isFunction, object} = require 'underscore'
escodegen = require 'escodegen'

vbParser = require './vb.parser'
exprParser = require './expr.parser'

repr = (arg) -> require('util').format '%j', arg
pprint = (arg) -> console.log require('util').inspect arg, false, null

member = (op) -> {'.': 'dot', '!': 'bang'}[op]
operator = (op) ->
  {'=': '===', '<>': '!==', '><': '!==', 'Or': '||', 'And': '&&'}[op] ? op

commonNodeValue = (n) ->
  switch n.name
    when '#document', 'source' # language.js nodes
      n.children?[0]?.value
    when 'expression'
      n.children[0].value
    when 'concat_expr'
      result = if n.children[1]? then literal '' # force string
      for {value}, i in n.children by 2
        result = if result? then binary '+', result, value else value
      result
    when 'add_expr', 'mul_expr'
      result = n.children[0].value
      for {value}, i in n.children by 2 when i > 0
        result = binary n.children[i-1]?.value, result, value
      result
    when 'mul_op', 'add_op', 'CMP_OP', 'AND', 'OR'
      n.innerText().replace /(\s|_)+$/, ''
    when 'start', 'value'
      n.children[0].value
    when 'bracketed_identifier'
      n.children[1].value
    when 'name_itself'
      n.innerText().replace /(\s|_)+$/, ''
    when 'name_in_brackets', 'lazy_name'
      n.innerText()
    when 'literal'
      literal n.children[1].value
    when 'literal_text'
      n.innerText()
    when 'identifier_op'
      n.innerText()
    when 'number'
      literal parseInt(n.innerText(), 10)
    when 'float'
      literal parseFloat(n.innerText())
    when 'like_expr'
      if n.children[2]? # /regexp/.test('string')
        memberCall literal(new RegExp n.children[2].value.value),
                   'test',
                   n.children[0].value
      else
        n.children[0].value
    when 'primary_expr'
      v = n.children[0].value

      # add .get() to finish up .dot(...), .bang(...), (...) expressions
      if v?.type == 'CallExpression' \
            #and not isEqual(v.callee, identifier 'ns') \
            and not (v.callee.type == 'MemberExpression' \
                     and isEqual(v.callee.property, identifier 'get') \
                     and v.arguments.length == 0)
        #console.log 'PRIMARY_EXPR', n.children[0].innerText(), v, 'YES'
        memberCall v, 'get'
      else
        #console.log 'PRIMARY_EXPR', n.children[0].innerText(), v, 'NO'
        v
    when 'name'
      n.children[0].value
    when 'braced_expression'
      n.children[1].value
    when 'TRUE'
      literal true
    when 'FALSE'
      literal false

vbNodeValue = (n) ->
  switch n.name
    when 'or_expr', 'and_expr'
      result = n.children[0].value
      for {value}, i in n.children by 2 when i > 0
        result = binary operator(n.children[i-1].value), result, value,
                        'LogicalExpression'
      result
    when 'cmp_expr'
      result = n.children[0].value
      if (right = n.children[2]?.value)?
        result = binary operator(n.children[1].value), result, right
      result
    when 'module'
      n.children[2].value
    when 'func_defs'
      type: 'ObjectExpression'
      properties: (value for {value} in n.children ? [])
    when 'func_def'
      [_1, name, args, body] = n.children

      # HACK traverse AST and replace all of `ns.get('arg').get()` with `arg`
      argnames = for a in args.value then a.name
      traverse = (obj) ->
        for k, v of obj when typeof v isnt 'string'
          found = no
          for argname in argnames
            ns_get_arg = memberCall identifier('ns'), 'get', literal argname
            if isEqual v, memberCall ns_get_arg, 'get'
              obj[k] = identifier argname
              found = yes
              break
          traverse v unless found
      traverse body.value

      type: 'Property'
      key:
        type: 'Literal'
        value: name.value
      value:
        type: 'FunctionExpression'
        id: null
        params: args.value
        defaults: []
        body: body.value
        rest: null
        generator: false
        expression: false
      kind: 'init'
    when 'args_spec'
      [l, args..., r, _1, _2] = n.children
      for arg in args by 2 then arg.value
    when 'arg_spec'
      identifier n.children[1].value
    when 'statements'
      type: 'BlockStatement'
      body: for {value} in n.children when value? then value
    when 'single_line_statement', 'multiline_statement', 'statement'
      n.children[0].value
    when 'exit_statement'
      type: 'ReturnStatement'
      argument: null
    when 'call_statement'
      type: 'ExpressionStatement'
      expression: call n.children[0].value, n.children[1].value
    when 'callee'
      n.children[0].value
    when 'argument_list'
      for {value} in n.children by 2 then value
    when 'positional_argument'
      n.children[0]?.value ? {type: 'Identifier', name: 'undefined'}
    when 'not_expr'
      if n.children[1]?
        type: 'UnaryExpression'
        operator: '!'
        argument: n.children[1].value
      else
        n.children[0].value
    when 'unrestricted_name'
      n.children[0].value
    when 'l_expression'
      n.children[0].value
    when 'name_expression', 'callee_name_expression'
      result = memberCall identifier('ns'), 'get', literal n.children[0].value
      for {value: operation} in n.children[1..]
        result = operation result
      result
    when 'member'
      (obj) ->
        memberCall obj,
                   member(n.children[0].value),
                   literal n.children[1].value
    when 'index'
      (callee) -> call callee, n.children[1].value
    when 'test_block'
      n.children[0].value
    when 'single_line_if_statement'
      [_1, test, _2, consequent] = n.children
      ifStatement test.value, consequent.value
    when 'if_statement'
      [_1, test, consequent, alternates..., _2] = n.children
      for {value: expression} in alternates[..].reverse()
        result = expression result
      ifStatement test.value, consequent.value, result
    when 'else_if_block'
      (alternate) -> ifStatement n.children[1].value,
                                 n.children[n.children.length-1].value,
                                 alternate
    when 'else_block'
      -> n.children[n.children.length-1].value
    when 'assign_statement'
      type: 'ExpressionStatement'
      expression:
        # Compile to `left`.let(`right`). Why not `left` = `right`? Because
        # JavaScript cannot assign to calls, while VBA can: a.b("c") = 42
        memberCall n.children[0].value, 'let', n.children[2].value

exprNodeValue = (n) ->
  switch n.name
    when 'identifier'
      literal n.children[0].value
    when 'identifier_expr', 'identifier_expr_part'
      n.children[0].value
    when 'identifier_expr_itself'
      # [A].[B]![C] => me('A').dot('B').bang('C')
      # About bang ! operator semantics:
      #   * http://stackoverflow.com/q/4804947
      #   * http://stackoverflow.com/q/2923957
      #   * http://www.cpearson.com/excel/DefaultMember.aspx
      #   * [MS-VBAL] 5.6.14 Dictionary Access Expressions
      result = n.children[0].value
      if result.type is 'Literal'
        result = memberCall identifier('me'), 'dot', result
      for {value: arg}, i in n.children by 2 when i > 0
        result = memberCall result, member(n.children[i-1].value), arg
      result
    when 'plain_call_expr'
      [{value: fn}, l, params..., r] = n.children
      nsCall(fn, for {value} in params by 2 then value)
    when 'lazy_call_expr'
      [{value: fn}, l, params..., r] = n.children
      nsCall(fn, for {value} in params by 2 then literal value)
    when 'lazy_name', 'lazy_value'
      n.innerText()

# if (`test`) { `consequent` } else { `alternate` }
ifStatement = (test, consequent, alternate = null) ->
  {type: 'IfStatement', test, consequent, alternate}

# `left` `operator` `right`
binary = (operator, left, right, type = 'BinaryExpression') ->
  {type, operator, left, right}

# `callee`(`args`...)
call = (callee, args) -> {type: 'CallExpression', callee, arguments: args}

# ns.get("`func`")(`args`...)
nsCall = (func, args) ->
  call memberCall(identifier('ns'), 'get', literal func), args

# <obj>.<property>(<args>...)
memberCall = (obj, property, args...) ->
  type: 'CallExpression'
  callee:
    type: 'MemberExpression'
    computed: no
    object: obj
    property: identifier property
  arguments: args

literal = (value) -> type: 'Literal', value: value
identifier = (name) -> type: 'Identifier', name: name

# Parse VB expression or module, and return Parser API AST [1] for escodegen.
#
# Escodegen [2] will translate this AST to JavaScript. To figure out how to
# construct a certain JavaScript fragment, use parser demo [3].
#
# [1]: https://developer.mozilla.org/en/SpiderMonkey/Parser_API
# [2]: https://github.com/Constellation/escodegen
# [3]: http://esprima.org/demo/parse.html
parse = (sourceType, expr) ->
  parser = {'vb': vbParser, 'expr': exprParser}[sourceType]
  tree = parser.parse expr

  # first copy-pasted from sqld3/parse_sql.coffee
  Object.getPrototypeOf(tree).toString = (spaces = '') ->
    result = try
               "=> #{escodegen.generate @value}"
             catch error
               if @value? then "=> AST: #{repr @value}" else ''

    string = spaces + "#{@name} <#{repr @innerText()}> " + result
    for child in @children when typeof child isnt 'string'
      string += "\n" + child.toString(spaces + ' ')

    return string

  nodeValue = {'vb': vbNodeValue, 'expr': exprNodeValue}[sourceType]
  tree.traverse
    traversesTextNodes: false
    exitedNode: (n) ->
      n.value = nodeValue(n) ? commonNodeValue(n)
      #if n.name is 'start' then console.log n.toString()

  if not tree.value? and process?.env?.TESTING?
    require("./test/#{sourceType}.peg.js").check '<string>', expr

  #pprint tree
  tree.value

compileExpression = (expr) ->
  tree = parse 'expr', expr
  unless tree?
    return 'Error parsing ' + expr
  #console.log 'TREE:'; pprint tree
  js = "var me = ns.get('Me'); return #{escodegen.generate tree};"
  #console.log 'JS: ', "`" + js + "`"
  js

compileModule = (code) ->
  tree = parse 'vb', code
  unless tree?
    throw "Error parsing module '#{code[..150]}...'"
  #console.log 'TREE:'; pprint tree
  js = "return #{escodegen.generate tree};"
  #console.log 'JS: ', "`" + js + "`"
  js

# compile VBA expression/module and run in {ns: ns} context
# TODO runExpression: separate `[bla]` from `bla`
runExpression = (expr, ns) -> runJS compileExpression(expr), ns
runModule = (code, ns) -> runJS compileModule(code), ns

# run JavaScript from string `js` in {ns: ns} context
# `ns` must be an object with `VBObject` interface
runJS = (js, ns) -> evaluate js, ns: ns

# better `eval`
evaluate = (js, context) ->
  keys = for key, val of context then key
  vals = for key, val of context then val
  try
    f = new Function keys..., js
  catch error
    console.log "#{error} in `#{js}`"
    throw error
  f vals...

class VBRuntimeError extends Error
  constructor: (msg) ->
    @name = 'VBRuntimeError'
    @message = msg or @name

# `VBObject` wraps `{attrs, default, type}` objects.
#
# It gives error-catching, case-insensitive interface to underlying `attrs`
# via .dot(), .get(), .let() methods. E.g.:
#
#     o = new VBObject
#               type: 'TextBox'
#               attrs: {visible: true, value: 'foo'}
#               default: 'value'
#
#     o.dot('visible').get() # => true
#     o.get('Visible') # => true
#     o.get('value') # => 'foo'
#     o.get() # => 'foo'
#
#     o.let('value', 'bar')
#     o.get() # => bar
#
class VBObject
  # Use `dot` argument for tests and debugging only
  constructor: ({attrs, @default, @type, @bang, dot}) ->
    if dot?
      @dot = dot
    else
      # lowercase keys
      @attrs = object([k.toLowerCase(), v] for k, v of attrs)

    # Fall back to `dot` if no `bang` function is provided. It's not 100%
    # correct, but most of the time `.` and `!` are indeed the same in VB.
    @bang ?= @dot

  dot: (attr) ->
    new Attribute this, attr

  get: (attr = @default) ->
    @attrs[@_lower attr]

  let: (attr, value) ->
    @attrs[@_lower attr] = value

  _lower: (attr) ->
    lower = attr.toLowerCase()
    unless @attrs[lower]?
      throw new VBRuntimeError "#{@type} has no attribute '#{attr}'"
    lower

# `Attribute` is used to defer `.dot(attr)` lookups in VBObject instances.
# TODO use it to defer `.bang()` lookups too
class Attribute
  constructor: (@object, @attr) ->

  dot: (attr) ->
    @object.get(@attr).dot(attr)

  bang: (attr) ->
    @object.get(@attr).bang(attr)

  get: (attr) ->
    if attr?
      @object.get(@attr).get(attr)
    else
      @object.get(@attr)

  let: (value) ->
    @object.let(@attr, value)

module.exports = {compileModule, compileExpression, runModule, runExpression,
                  VBObject, evaluate, VBRuntimeError}

# Usage: cat VBA_module | coffee vb.coffee
#        echo -n "[foo]&[bar]" | coffee vb.coffee -e
if require.main == module
  process.stdin.resume()
  process.stdin.setEncoding 'utf8'

  data = ''
  process.stdin.on 'data', (chunk) -> data += chunk
  process.stdin.on 'end', ->
    c = if process.argv[2] is '-e' then compileExpression else compileModule
    process.stdout.write c(data) + '\n'
