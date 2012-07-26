parser = require "./vb.js"
escodegen = require "escodegen"

exports.evaluate = (expr, Me) ->
    (compile parse expr)(Me)

repr = (arg) -> require('util').format '%j', arg
pprint = (arg) -> console.log require('util').inspect arg, false, null

parse = (expr) ->
    tree = parser.parse expr

    # copy-pasted from sqld3/parse_sql.coffee
    Object.getPrototypeOf(tree).toString = (spaces) ->
        if not spaces then spaces = ""

        value = (if this.value? then "=> #{repr this.value}" else '')
        string = spaces + this.name +  " <" + this.innerText() + "> " + value
        children = this.children
        index = 0

        for child in children
            if typeof child == "string"
                #string += "\n" + spaces + ' ' + child
            else
                string += "\n" + child.toString(spaces + ' ')

        return string

    tree.traverse
        traversesTextNodes: false
        exitedNode: (n) ->
            n.value = switch n.name
                when 'source', '#document' # language.js nodes
                    n.children[0].value
                when 'start', 'expression', 'value'
                    n.children[0].value
                when 'literal'
                    type: 'Literal'
                    value: n.children[1].value
                when 'literal_text'
                    n.innerText()
                when 'identifier_expr'
                    result =
                        type: 'Identifier'
                        name: 'Me'
                    for {value}, i in n.children by 2
                        result = switch op ? '.'
                            when '.' # A.B => A[B]
                                type: 'MemberExpression'
                                computed: yes
                                object: result
                                property: value
                            when '!'
                                # bang op: A!B => A.__default(B), links:
                                # * http://stackoverflow.com/q/4804947
                                # * http://stackoverflow.com/q/2923957
                                # * http://www.cpearson.com/excel/DefaultMember.aspx
                                type: 'CallExpression'
                                callee: 
                                    type: 'MemberExpression'
                                    computed: no
                                    object: result
                                    property:
                                        type: 'Identifier'
                                        name: '__default'
                                'arguments': [ value ]
                        op = n.children[i+1]?.value
                    result
                when 'identifier_op'
                    n.innerText()
                when 'identifier'
                    type: 'Literal'
                    value: n.children[1].value
                when 'name'
                    n.innerText()
                when 'add_expr' 
                    result = if n.children[1]?.name is 'concat_op'
                                 type: 'Literal', value: '' # force string
                    for {value}, i in n.children by 2
                        result = if result? then plus result, value else value
                    result

            #if n.name is 'start' then console.log n.toString()

    tree.value

plus = (left, right) ->
    type: 'BinaryExpression'
    operator: '+'
    left: left
    right: right

# compile to JavaScript
compile = (tree) ->
    #console.log 'TREE:'
    #pprint tree
    code = escodegen.generate tree
    #console.log 'CODE', code
    new Function 'Me', "return #{code};"

# Usage: coffee vbjs.coffee "[foo]&[bar]"
#parse process.argv[2]
