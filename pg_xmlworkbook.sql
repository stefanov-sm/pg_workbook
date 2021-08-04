----------------------------------------
-- pg_xmlworkbook, S. Stefanov, Aug-2021
----------------------------------------

CREATE OR REPLACE FUNCTION pg_xmlworkbook (arg_queries_array json, arg_sheet_names_array json default '[]', arg_parameters_array json default '[]')
RETURNS SETOF text LANGUAGE plpgsql AS
$function$

DECLARE
WORKBOOK_HEADER constant text[] := array
[
'<?xml version="1.0" encoding="utf8"?>',
'<?mso-application progid="Excel.Sheet"?>',
'<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">',
'  <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">',
'   <Subject>Postgres spreadsheet export</Subject>',
'   <Author>pg_xmlspreadsheet</Author>',
'   <Company>https://github.com/stefanov-sm/pg_xmlspreadsheet</Company>',
'  </DocumentProperties>',
'  <Styles>',
'   <Style ss:ID="Default" ss:Name="Normal"><Font ss:FontName="Arial" ss:Size="10" ss:Color="#000000"/></Style>',
'   <Style ss:ID="Href" ss:Name="Hyperlink"><Font ss:FontName="Arial" ss:Size="10" ss:Color="#0000FF" ss:Underline="Single"/></Style>',
'   <Style ss:ID="Date"><NumberFormat ss:Format="Short Date"/></Style>',
'   <Style ss:ID="DateTime"><NumberFormat ss:Format="yyyy-mm-dd hh:mm:ss"/></Style>',
'   <Style ss:ID="Header">',
'    <Borders>',
'     <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>',
'     <Border ss:Position="Top"    ss:LineStyle="Continuous" ss:Weight="1"/>',
'     <Border ss:Position="Left"   ss:LineStyle="Continuous" ss:Weight="1"/>',
'     <Border ss:Position="Right"  ss:LineStyle="Continuous" ss:Weight="1"/>',
'    </Borders>',
'    <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>',
'   </Style>',
'  </Styles>'
];

WORKSHEET_HEADER constant text[] := array
[
'  <Worksheet ss:Name="__VALUE__">',
'  <Table>'
];

WORKSHEET_FOOTER constant text[] := array
[
'</Table>',
'  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">',
'   <FreezePanes/><FrozenNoSplit/><SplitHorizontal>1</SplitHorizontal>',
'   <TopRowBottomPane>1</TopRowBottomPane><ActivePane>2</ActivePane>',
'  </WorksheetOptions>',
'  </Worksheet>'
];

WORKBOOK_FOOTER constant text := '</Workbook>';

TITLE_ITEM    constant text := '    <Cell ss:StyleID="Header"><Data ss:Type="String">__VALUE__</Data></Cell>';
DATE_ITEM     constant text := '    <Cell ss:StyleID="Date"><Data ss:Type="DateTime">__VALUE__</Data></Cell>';
DTIME_ITEM    constant text := '    <Cell ss:StyleID="DateTime"><Data ss:Type="DateTime">__VALUE__</Data></Cell>';
HREF_ITEM     constant text := '    <Cell ss:StyleID="Href" ss:HRef="__HREF__"><Data ss:Type="String">__VALUE__</Data></Cell>';
TEXT_ITEM     constant text := '    <Cell><Data ss:Type="String">__VALUE__</Data></Cell>';
NUMBER_ITEM   constant text := '    <Cell><Data ss:Type="Number">__VALUE__</Data></Cell>';
BOOL_ITEM     constant text := '    <Cell><Data ss:Type="Boolean">__VALUE__</Data></Cell>';
EMPTY_ITEM    constant text := '    <Cell></Cell>';
COLUMN_ITEM   constant text := '   <Column ss:AutoFitWidth="0" ss:Width="__VALUE__"/>';
BEGIN_ROW     constant text := '   <Row>';
END_ROW       constant text := '   </Row>';

SR_TOKEN      constant text := '__VALUE__';
HREF_TOKEN    constant text := '__HREF__';
HREF_REGEX    constant text := '^#(.+)##(.+)';

AVG_CHARWIDTH constant integer := 5.5;
MIN_FLDWIDTH  constant integer := 40;
TS_CHOP_SIZE  constant integer := 19;

r record;
jr json;
v_key text;
v_value text;
href_array text[];
column_types text[];
running_line text;
running_column integer;
cold boolean;

running_query text;
running_index integer;
running_param json;

BEGIN

foreach running_line in array WORKBOOK_HEADER loop
  return next running_line;
end loop;

for running_query, running_index in (select * from json_array_elements_text(arg_queries_array) with ordinality) loop

  running_param := coalesce(arg_parameters_array -> (running_index - 1), '{}'::json);
  return next replace(WORKSHEET_HEADER[1], SR_TOKEN, coalesce(arg_sheet_names_array ->> (running_index - 1), format('Sheet%s', running_index)));
  return next WORKSHEET_HEADER[2];
  cold := true;

  for r in execute dynsql_safe(running_query, running_param) using running_param loop

    jr := to_json(r);
    if cold then
      column_types := (select array_agg(json_typeofx("value")) from json_each(jr));
      for v_key in select "key" from json_each_text(jr) loop
        running_line := replace(COLUMN_ITEM, SR_TOKEN, greatest(length(v_key) * AVG_CHARWIDTH, MIN_FLDWIDTH)::text);
        return next running_line;
      end loop;
      return next BEGIN_ROW;
      for v_key in select "key" from json_each_text(jr) loop
        running_line := replace(TITLE_ITEM, SR_TOKEN, xml_escape(v_key));
        return next running_line;
      end loop;
      return next END_ROW;
      cold := false;
    end if;

    return next BEGIN_ROW;

    for v_key, v_value, running_column in (select * from json_each_text(jr) with ordinality) loop
      if v_value is null then
        running_line := EMPTY_ITEM;
      else
        if column_types[running_column] = 'null' then
          column_types[running_column] := json_typeofx(jr -> v_key);
        end if;
        case column_types[running_column]
          when 'string'   then running_line := replace(TEXT_ITEM,   SR_TOKEN, xml_escape(v_value));
          when 'number'   then running_line := replace(NUMBER_ITEM, SR_TOKEN, v_value);
          when 'boolean'  then running_line := replace(BOOL_ITEM,   SR_TOKEN, v_value::boolean::int::text);
          when 'date'     then running_line := replace(DATE_ITEM,   SR_TOKEN, v_value);
          when 'datetime' then running_line := replace(DTIME_ITEM,  SR_TOKEN, left(v_value, TS_CHOP_SIZE));
          when 'href'     then href_array   := regexp_matches(xml_escape(v_value), HREF_REGEX);
                               running_line := replace(replace(HREF_ITEM, SR_TOKEN, href_array[1]), HREF_TOKEN, href_array[2]); 
          else                 running_line := replace(TEXT_ITEM,   SR_TOKEN, xml_escape(v_value));
        end case;
      end if;
      return next running_line;
    end loop;
    return next END_ROW;
  end loop;

  foreach running_line in array WORKSHEET_FOOTER loop
    return next running_line;
  end loop;

end loop;

return next WORKBOOK_FOOTER;

END;
$function$;
