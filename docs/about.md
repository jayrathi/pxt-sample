# About Sample Target

Example of editor target for [Microsoft MakeCode](https://makecode.com/).

See [GitHub repo](https://github.com/Microsoft/pxt-sample) for details.

# Questions to be answered

1. How does the editor store the block files? Are they converted to JS/TS? Are the blocks created back from JS/TS?

2. 

# Understanding of the architecture -

https://makecode.com/target-creation

The sim folder contains the important code pieces.
* api.ts has the definition of the api. It is compiled into libs\core\sims.d.ts. The editor creates the blocks from this d.ts.

* Important links

Editor extension sample https://github.com/samelhusseini/pxt-editor-extension-sample

