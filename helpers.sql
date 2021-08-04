---------------------------------------------------------------
-- pg_xmlworkbook helpers, S. Stefanov, Luca Ferrari, July-2020
---------------------------------------------------------------

/**
  * Performs tag escaping.
  *
  * Example:
    testdb=> select xml_escape( '<xml version="1.0"> <equation> 10 > 20 && 20 < 100 </equation> </xml>' );
                                                   xml_escape
  -------------------------------------------------------------------------------------------------------------
   &lt;xml version="1.0"&gt; &lt;equation&gt; 10 &gt; 20 &amp;&amp; 20 &lt; 100 &lt;/equation&gt; &lt;/xml&gt;
  (1 row)
*/
create or replace function xml_escape(s text)
returns text language sql immutable strict as
$$
    select  replace( replace( replace( s, '&', '&amp;' )
                                , '>'
                                , '&gt;' )
                            , '<'
                            , '&lt;' );
$$;

/**
  * Performs injection-safe macro expansion.
  * A macro is an upper case K&R identifier with a double underscore at the beginning and at the end,
  * like for example __FOO__.
  * Macros are globally substituted with json attribute text expressions from `args` json(b) object.
  * `args` attribute names are restricted to K&R identifiers. 
  * Valuable quality features by Luca Ferrari
  * 
  * Intended Use: EXECUTE dynsql_safe(sql_template, json_args) USING json_args;
  * 
  * Examples:

  select dynsql_safe
  (
   'select x from t where y = __A__::integer and z <> __B__;', 
   '{"a":1, "b":"two", "good_one":"three"}'
  );
  ------------------------------------
  select x from t where y = ($1->>'a')::integer and z <> ($1->>'b');
  
  select dynsql_safe
  (
   'select x from t where y = __A__::integer and z <> __B__;', 
   '{"a":1, "b":"two", "bad one":"three"}'
  );
  ------------------------------------
  SQL Error [P0001]: ERROR: Non-K&R key found in JSON(B) arguments
    Hint: Offending key: "bad one"
    Where: PL/pgSQL function dynsql_safe(text,jsonb) line 11 at RAISE
  
  select dynsql_safe
  (
   'select x from t where y = __A__::integer and z <> __B__;', 
   '{"a":1, "bb":"two"}'
  );
  ------------------------------------
  SQL Error [P0001]: ERROR: 1 macro(s) not processed, please check your JSON(B) arguments!
    Hint: Macro(s) left: __B__
    Where: PL/pgSQL function dynsql_safe(text,jsonb) line 19 at RAISE
*/
create or replace function dynsql_safe(arg_query text, args json) returns text as
$$
declare
    KRUCID_RX constant text := '[_A-Z][_A-Z0-9]*';
    running_key text;
    tokens_left text[];
begin
    -- Rewrite the query. Convert __MACRO__ placeholders into json attribute text expressions
    for running_key in select "key" from json_each_text(args) loop
        if running_key ~* ('^'||KRUCID_RX||'$') then
            arg_query := replace(arg_query, '__'||upper(running_key)||'__', '($1->>'''||running_key||''')');
        else
            raise exception 'Non-K&R key found in JSON(B) arguments'
            using hint = 'Offending key: "'||running_key||'"';
        end if;
    end loop;

    -- check there is no macro without expansion
    tokens_left := array(select regexp_matches(arg_query, '__'||KRUCID_RX||'__', 'g')); -- upper case only
    if array_length(tokens_left, 1) > 0 then
        raise exception '% macro(s) not processed, please check your JSON(B) arguments!', array_length(tokens_left, 1)
        using hint = 'Macro(s) left: '||array_to_string(tokens_left, ', ');
    end if;             
    return arg_query;
end;
$$ language plpgsql;

/**
  * json_typeof extension
  * detects date, datetime and href types.
*/
create or replace function json_typeofx(j json)
returns text language sql immutable strict as
$$
select
  case
    when json_typeof(j) = 'string' then case 
      when j::text ~ '^"\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d(\.\d+)?' then 'datetime'
      when j::text ~ '^"\d{4}-\d\d-\d\d' then 'date'
      when j::text ~ '^"#.+##.+' then 'href'
      else 'string' 
    end
    else json_typeof(j)
  end;
$$;
