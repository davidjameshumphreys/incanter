(ns experimental.update
 (:use [incanter core charts]))

(comment (def e1 (java.util.concurrent.Executors/newSingleThreadScheduledExecutor))
(def e2 (java.util.concurrent.Executors/newSingleThreadScheduledExecutor))


(def sched (. e1 scheduleAtFixedRate (fn [] (prn 3)) 0 5 java.util.concurrent.TimeUnit/SECONDS))

(prn (type sched))

(. e2 schedule
 (proxy [java.util.concurrent.Callable] [](call [] (do (prn "closing")(. sched cancel true))))
 (long 5)
 java.util.concurrent.TimeUnit/SECONDS)
(def d (dataset ["a" "b"] [[1 2] [3 4]]))

(defn appendable-dataset [#^dataset ds]
  (let [dc (:column-names ds)
        dr (:rows ds)]
  (ref (dataset dc (vec dr)))))

(def dd (appendable-dataset d))

(defn alter-d [ds new-row]
  (dataset (:column-names ds)
  (vec
    (if (nil? new-row)
      (:rows ds)
      (conj (:rows ds)
        (apply assoc {} (interleave (:column-names ds) new-row)))))))

(dosync (alter dd alter-d [5 6]))
(dosync (alter dd alter-d [7 8]))
(dosync (alter dd alter-d [9 0])))
