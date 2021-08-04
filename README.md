# pg_xmlworkbook
Export results of several Postgresql queries as a multi-sheet SpreadsheetML workbook (XML spreadsheet format).  
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
select xline from util.pg_xmlworkbook
(
    json '["select A ...", "select B ...", "select C ..."]', -- arg_queries_array
    json '["Sheet-A", "Sheet-B", "Sheet-C"]',                -- arg_sheet_names_array
    json '[                                                  -- arg_parameters_array
           {"argument_1":1, "argument_2":"A"}, 
           {"argument_11":10, "argument_12":"B"}, 
           {"argument_21":100, "argument_22":"Z"}
          ]'
) xline; 
```
### Example. Create a three-sheet workbook out of three trivial parameterized queries 
```PGSQL
COPY
(
  select * from util.pg_xmlworkbook
  (
    to_json(array[
      $query_a$
      select
        v as "value",
        to_char(v % 4000, 'FMRN') as "mod 4000 roman",
        v^2 as "square",
        v^3 as "cube",
        clock_timestamp() as "date and time",
        '#<see more>##https://www.google.com/search?q='||v::text as "search Google"
      from generate_series(__FROM__::integer, __TO__::integer, 1) t(v)
      where v::text like __PATTERN__;
      $query_a$,
      $query_b$
      select
        v as "Стойност",
        to_char(v % 4000, 'FMRN') as "Римски цифри, mod 4000",
        v^2 as "На квадрат",
        v^3 as "На трета степен",
        clock_timestamp() as "Дата & час",
        '#<Виж повече>##https://www.google.com/search?q='||v::text as "Потърси го в Google"
      from generate_series(__FROM__::integer, __TO__::integer, 1) t(v)
      where v::text like __PATTERN__;
      $query_b$,
      $query_c$
      select
        v as "Native value",
        to_char(v % 4000, 'FMRN') as "SPQR",
        v^2 as "Square value",
        v^3 as "Cube value",
        clock_timestamp() as "Event date & time",
        '#<Search the web>##https://www.google.com/search?q='||v::text as "Google it now!"
      from generate_series(__ABC__::integer, __XYZ__::integer, 1) t(v)
      where v::text ~ __RX__;
      $query_c$
    ]::text[]),
    json '["Threes", "Четворки", "Fives"]',
    json '[
        {"from":10,  "to":1010,  "pattern":"%3%"},
        {"from":100, "to":1100,  "pattern":"%4%"}, 
        {"abc":1000, "xyz":2000, "rx":"5"}
    ]'
  )
) TO '/path/to/delme.xml';
```
