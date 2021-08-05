# pg_xmlworkbook
Export results of Postgresql queries as a multi-sheet SpreadsheetML workbook (XML spreadsheet format).  
The prototype of the function is as follows:
```PGSQL
FUNCTION pg_xmlworkbook
(
  arg_queries_array json, 
  arg_sheet_names_array json default '[]', 
  arg_parameters_array json default '[]'
)
RETURNS SETOF text;
```
### How to use
```PGSQL
select xline from pg_xmlworkbook
(
    json '["select A ...", "select B ...", "select C ..."]', -- arg_queries_array
    json '["Sheet-A", "Sheet-B", "Sheet-C"]',                -- arg_sheet_names_array
    json '[
           {"argument_1":1, "argument_2":"A"}, 
           {"argument_11":10, "argument_12":"B"}, 
           {"argument_21":100, "argument_22":"Z"}
          ]'                                                 -- arg_parameters_array
) xline; 
```
__arg_queries_array__ is an array of SQL queries that may be parameterized.  
Parameter placeholders are defined as valid uppercase identifiers with two underscores as prefix and suffix.  
`__FROM__`, `__TO__`, `__PATTERN__`   

__NB__: Placeholders are rewritten into runtime expressions that _always_ return type `text`. This is why they may need to be explicitly cast.  
`__FROM__::integer, __TO__::integer` in the example below  
  
Optional __arg_parameters_array__ is an array of JSON objects with parameters' names/values.  
  `{"from":15, "to":100015, "pattern":"%3%"}`  

Parameter names are K&R case-insesnitive identifiers.  
### Example. Create a three-sheet workbook out of three trivial parameterized queries 
```PGSQL
-- Postgres server-side queries to workbook of spreadsheet example
COPY
(
  select * from pg_xmlworkbook
  (
    to_json(array[
      $query_a$
        SELECT
          v AS "value",
          to_char(v % 4000, 'FMRN') AS "mod 4000 roman",
          v^2 AS "square",
          v^3 AS "cube",
          clock_timestamp() AS "date and time",
          format('#<see more about %1$s>##https://www.google.com/search?q=%1$s', v) AS "search Google"
        FROM generate_series(__FROM__::integer, __TO__::integer, 1) t(v)
        WHERE v::text LIKE __PATTERN__;
      $query_a$,
      $query_b$
        SELECT
          v AS "Стойност",
          to_char(v % 4000, 'FMRN') AS "Римски цифри, mod 4000",
          v^2 AS "На квадрат",
          v^3 AS "На трета степен",
          clock_timestamp() AS "Дата & час",
          format('#<Виж повече за %1$s>##https://www.google.com/search?q=%1$s', v) AS "Потърси го в Google"
        FROM generate_series(__FROM__::integer, __TO__::integer, 1) t(v)
        WHERE v::text LIKE __PATTERN__;
      $query_b$,
      $query_c$
        SELECT
          v AS "Native value",
          to_char(v % 4000, 'FMRN') AS "SPQR",
          v^2 AS "Square value",
          v^3 AS "Cube value",
          clock_timestamp() AS "Event date & time",
          format('#<Search the web for %1$s>##https://www.google.com/search?q=%1$s', v) AS "Google it now!"
        FROM generate_series(__ABC__::integer, __XYZ__::integer, 1) t(v)
        WHERE v::text ~ __RX__;
      $query_c$
    ]::text[]),
    json '["Threes", "Четворки", "Fives"]',
    json '[
           {"from":10,  "to":1010,  "pattern":"%3%"},
           {"from":100, "to":1100,  "pattern":"%4%"}, 
           {"abc":1000, "xyz":2000, "rx":"5"}
          ]'
  )
) TO 'd:/temp/delme.xml';
```
