# docx-utils
Library that helps to work with `.docx` files.

Essentially, it is a Clojure wrapper on [Apache POI](https://poi.apache.org/) library.  

The main idea is to use plain data instead of Apache POI classes to transform a `.docx` document.

There are 2 types of transformations: append and replace.

It is possible to append and replace the following data types:
- in-line text;
- paragraph text;
- image;
- table;
- bullet list;
- numbered list.

# Latest Version
The latest version of `docx-utils` is hosted on [Clojars](https://clojars.org/lt.tokenmill/docx-utils):

[![Clojars Project](https://img.shields.io/clojars/v/lt.tokenmill/docx-utils.svg)](https://clojars.org/lt.tokenmill/docx-utils)

# Main concepts
Main entrance function is `docx-utils.core/transform`. It has 3 versions:
- 1 argument version takes in a list of transformations, it creates an empty `.docx` document, applies all transformations and returns a file path of the resulting document. Note that Resulting file is created in a temp folder and it will be deleted after the JVM exits. This function could/should be used to create a template `docx` document.
- 2 argument version takes in a template-file-path and a list of transformations. It loads the template document applies transformations and returns the resulting document. If the template-file-path is nil then new `docx` is created and used as a template.
- 3 argument version takes in a template-file-path, output-file-path and a list of transformations. It loads a template document, applies the transformations and stores the resulting document in a file under the output-file-path. If output-file-path is `nil` then error is thrown.

Note that if a placeholder is not found in the template then nothing happens, only a log warning is produced.

# Documentation
The complete [API documentation](https://tokenmill.github.io/docx-utils/) is also available (codox generated).

# Examples
To replace a placeholder with a paragraph text:
```clojure
(docx-utils.core/transform "/path/to/template/file.docx"
                           [{:type        :replace-text
                             :placeholder "${PLACEHOLDER}"
                             :replacement "Standalone paragraph."}])
```

To replace a placeholder inside a text paragraph:
```clojure
(docx-utils.core/transform "/path/to/template/file.docx"
                           [{:type        :replace-text-inline
                             :placeholder "${PLACEHOLDER}"
                             :replacement "in-lined text"}])
```

To replace a placeholder with a data table:
```clojure
(docx-utils.core/transform "/path/to/template/file.docx"
                           [{:type        :replace-table
                             :placeholder "${PLACEHOLDER}"
                             :replacement [["cell 11" "cell 12" "cell 13"] ["cell 21" "cell 22" "cell 23"]]}])
```

To replace a placeholder with a bulleted list:
```clojure
(docx-utils.core/transform "/path/to/template/file.docx"
                           [{:type        :replace-bullet-list
                             :placeholder "${PLACEHOLDER}"
                             :replacement ["item 1" "item 2" "item 3"]}])
```

To replace a placeholder with a numbered list:
```clojure
(docx-utils.core/transform "/path/to/template/file.docx"
                           [{:type        :replace-numbered-list
                             :placeholder "${PLACEHOLDER}"
                             :replacement ["item 1" "item 2" "item 3"]}])

To replace a placeholder with an image:
```clojure
(docx-utils.core/transform "/path/to/template/file.docx"
                           [{:type        :replace-image
                             :placeholder "${PLACEHOLDER}"
                             :replacement "/path/to/image/file.jpg"}])
```

To append a text paragraph to the end of the template document:
```clojure
(docx-utils.core/transform "/path/to/template/file.docx"
                           [{:type        :append-text
                             :replacement "Text paragraph."}])
```

To append a text snippet to the end of the last paragraph of the template document:
```clojure
(docx-utils.core/transform "/path/to/template/file.docx"
                           [{:type        :append-text-inline
                             :replacement "text snippet"}])
```

To append a data table to the end of the template document:
```clojure
(docx-utils.core/transform "/path/to/template/file.docx"
                           [{:type        :append-table
                             :replacement [["cell 11" "cal 12" "cell 13"] ["cell 21" "cell 22" "cell 23"]]}])
```

To append a bullet list to the end of the template document:
```clojure
(docx-utils.core/transform "/path/to/template/file.docx"
                           [{:type        :append-bullet-list
                             :replacement ["item a" "item b" "item c"]}])
```

To append a numbered list to the end of the template document:
```clojure
(docx-utils.core/transform "/path/to/template/file.docx"
                           [{:type        :append-numbered-list
                             :replacement ["item a" "item b" "item c"]}])
```

To append an image to the end of the template document:
```clojure
(docx-utils.core/transform "/path/to/template/file.docx"
                           [{:type        :replace-image
                             :replacement "/path/to/image/file.jpg"}])

# Future work
- `:replacement` could be either a `String` or a map. If `String` then the value is pasted into the document without any additional formating (the formating of the placeholder is preserved), if a map is provided the underlying `Run` is formated accordingly to the options provided in a map, e.g. {:bold true :text "Bolded text"}.
- `:type` value should be somehow taken from a list of values. Maybe a namespace with constants? e.g. `(def REPLACE_TEXT :replace-text)`. Or Java enums?
- expose a Java interface to the library.
