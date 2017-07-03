(defproject docx-utils "0.1.0-SNAPSHOT"
  :description "FIXME: write description"
  :url "http://example.com/FIXME"
  :license {:name "Eclipse Public License"
            :url "http://www.eclipse.org/legal/epl-v10.html"}
  :dependencies [[org.clojure/clojure "1.8.0"]
                 [org.clojure/tools.logging "0.3.1"]
                 [org.apache.xmlbeans/xmlbeans "2.6.0"]
                 [org.apache.poi/poi "3.16"]
                 [org.apache.poi/poi-ooxml "3.16"]
                 [org.apache.poi/ooxml-schemas "1.3"]
                 [org.apache.poi/poi-ooxml-schemas "3.16"]]
  :resource-paths ["resources"]
  :profiles {:dev {:resource-paths ["test/resources"]}
             :repl {:resource-paths ["test/resources"]}})
