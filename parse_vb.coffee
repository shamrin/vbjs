# Usage: coffee parse_vb.coffee "[foo]&[bar]"

parser = require "./vb.js"

tree = parser.parse process.argv[2]

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
    exitedNode: (node) ->
        if node.name is 'start'
            console.log node.toString()
