escodegen = require "escodegen"
vb_parser = require "./vb.parser"
expr_parser = require "./expr.parser"

repr = (arg) -> require('util').format '%j', arg
pprint = (arg) -> console.log require('util').inspect arg, false, null

member = (op) -> {'.': 'dot', '!': 'bang'}[op]

common_node_value = (n) ->
  switch n.name
    when '#document', 'source' # language.js nodes
      n.children?[0]?.value
    when 'expression'
      n.children[0].value
    when 'concat_expr'
      result = if n.children[1]? then literal '' # force string
      for {value}, i in n.children by 2
        result = if result? then operate '+', result, value else value
      result
    when 'add_expr', 'mul_expr'
      for {value}, i in n.children by 2
        op = n.children[i-1]?.value
        result = if result? then operate op, result, value else value
      result
    when 'mul_op', 'add_op'
      n.innerText().replace /\s+$/, ''
    when 'start', 'value', 'identifier_expr', 'identifier_expr_part'
      n.children[0].value
    when 'identifier'
      literal n.children[0].value
    when 'bracketed_identifier'
      n.children[1].value
    when 'name_itself'
      n.innerText().replace /\s+$/, ''
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
      n.children[0].value # FIXME it's just a stub now
    when 'primary_expr'
      n.children[0].value
    when 'name'
      n.children[0].value
    when 'plain_call_expr'
      [{value: fn}, l, params..., r] = n.children
      call(fn, for {value} in params by 2 then value)
    when 'braced_expression'
      n.children[1].value

vb_node_value = (n) ->
  switch n.name
    when 'or_expr', 'and_expr', 'cmp_expr'
      n.children[0].value # FIXME it's just a stub now
    when 'module'
      n.children[2].value
    when 'func_defs'
      type: 'ObjectExpression'
      properties: (value for {value} in n.children ? [])
    when 'func_def'
      name = n.children[1].value
      body = n.children[3].value
      type: 'Property'
      key:
        type: 'Literal'
        value: name
      value:
        type: 'FunctionExpression'
        id: null
        params: []
        defaults: []
        body:
          type: 'BlockStatement'
          body: body
        rest: null
        generator: false
        expression: false
      kind: 'init'
    when 'statements'
      for {value} in n.children when value? then value
    when 'single_line_statement', 'multiline_statement'
      n.children[0].value
    when 'statement'
      n.children[0].value
    when 'exit_statement'
      type: 'ReturnStatement'
      argument: null
    when 'call_statement'
      type: 'ExpressionStatement'
      expression:
        type: 'CallExpression'
        callee: n.children[0].value
        arguments: n.children[1]?.value ? []
    when 'callee'
      n.children[0].value
    when 'argument_list'
      for {value} in n.children by 2 then value
    when 'positional_argument'
      n.children[0]?.value ? {type: 'Identifier', name: 'undefined'}
    when 'identifier_expr_itself'
      # [A].[B]![C] => ns('A').dot('B').bang('C')
      result = n.children[0].value
      if result.type is 'Literal'
        result =
          type: 'CallExpression'
          callee: identifier 'ns'
          arguments: [ result ]
      for {value: arg}, i in n.children by 2 when i > 0
        result =
          type: 'CallExpression'
          callee:
            type: 'MemberExpression'
            computed: no
            object: result
            property: identifier member n.children[i-1].value
          arguments: [ arg ]
      result
    when 'not_expr'
      n.children[1]?.value ? n.children[0].value  # FIXME it's just a stub now
    when 'unrestricted_name'
      n.children[0].value
    when 'name_expression', 'callee_name_expression'
      result =
        type: 'CallExpression'
        callee: identifier 'ns'
        arguments: [ literal n.children[0].value ]
      for {value: operation} in n.children[1..]
        result = operation result
      result
    when 'member'
      (object) ->
        type: 'CallExpression'
        callee:
          type: 'MemberExpression'
          computed: no
          object: object
          property: identifier member n.children[0].value
        arguments: [ literal n.children[1].value ]
    when 'index'
      [l, params..., r] = n.children
      (callee) ->
        type: 'CallExpression'
        callee: callee
        arguments: for {value} in params by 2 then literal value
    when 'test_block'
      n.children[0].value
    when 'if_statement'
      [_, test_block, then_block] = n.children
      type: 'IfStatement'
      test: test_block.value
      consequent:
        type: 'BlockStatement'
        body: then_block.value
      alternate: null

expr_node_value = (n) ->
  switch n.name
    when 'identifier_expr_itself'
      # [A].[B]![C] => me('A').dot('B').bang('C')
      # About bang ! operator semantics:
      #   * http://stackoverflow.com/q/4804947
      #   * http://stackoverflow.com/q/2923957
      #   * http://www.cpearson.com/excel/DefaultMember.aspx
      result = n.children[0].value
      if result.type is 'Literal'
        result =
          type: 'CallExpression'
          callee: identifier 'me'
          arguments: [ result ]
      for {value: arg}, i in n.children by 2 when i > 0
        result =
          type: 'CallExpression'
          callee:
            type: 'MemberExpression'
            computed: no
            object: result
            property: identifier member n.children[i-1].value
          arguments: [ arg ]
      result
    when 'lazy_call_expr'
      [{value: fn}, l, params..., r] = n.children
      call(fn, for {value} in params by 2 then literal value)
    when 'lazy_name'
      n.innerText()
    when 'lazy_value'
      n.innerText()

# Parse VB expression or module, and return Parser API AST [1] for escodegen.
#
# Escodegen [2] will translate this AST to JavaScript. To figure out how to
# construct a certain JavaScript fragment, use parser demo [3].
#
# [1]: https://developer.mozilla.org/en/SpiderMonkey/Parser_API
# [2]: https://github.com/Constellation/escodegen
# [3]: http://esprima.org/demo/parse.html
parse = (source_type, expr) ->
  parser = {'vb': vb_parser, 'expr': expr_parser}[source_type]
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

  node_value = {'vb': vb_node_value, 'expr': expr_node_value}[source_type]
  tree.traverse
    traversesTextNodes: false
    exitedNode: (n) ->
      n.value = node_value(n) ? common_node_value(n)
      #if n.name is 'start' then console.log n.toString()

  if not tree.value? and process?.env?.TESTING?
    require("./test/#{source_type}.peg.js").check '<string>', expr

  #pprint tree
  tree.value

# `left` `op` `right`
operate = (op, left, right) ->
  type: 'BinaryExpression'
  operator: op
  left: left
  right: right

# ns("`func_name`")(ns, `args`...)
call = (func_name, args) ->
  type: 'CallExpression'
  callee:
    type: 'CallExpression'
    callee: identifier 'ns'
    arguments: [literal func_name]
  arguments: [identifier('ns')].concat args

literal = (value) -> type: 'Literal', value: value
identifier = (name) -> type: 'Identifier', name: name

compileExpression = (expr) ->
  tree = parse 'expr', expr
  unless tree?
    return 'Error parsing ' + expr
  #console.log 'TREE:'; pprint tree
  js = "var me = ns('Me').dot; return #{escodegen.generate tree};"
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
runExpression = (expr, ns) -> runJS compileExpression(expr), ns
runModule = (code, ns) -> runJS compileModule(code), ns

# run JavaScript from string `js` in {ns: ns} context
runJS = (js, ns) ->
  evaluate js, ns: (name) ->
                     unless ns[name]?
                       throw new VBRuntimeError "VB name '#{name}' not found"
                     ns[name]

# better than `eval`
evaluate = (js, context) ->
  keys = for key, val of context then key
  vals = for key, val of context then val
  try
    f = new Function keys..., js
  catch error
    console.log "#{error} running `#{js}`"
    throw error
  f vals...

class VBRuntimeError extends Error
  constructor: (msg) ->
    @name = 'VBRuntimeError'
    @message = msg or @name

module.exports = {compileModule, compileExpression, runModule, runExpression,
                  evaluate, VBRuntimeError}

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
