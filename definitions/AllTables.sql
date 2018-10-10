SELECT table_name, tablespace_name, num_rows , last_analyzed , last_analyzed as last_analyzed_at
  FROM all_tables 
 WHERE tablespace_name is not null  and num_rows > :NumRows  order by num_rows desc