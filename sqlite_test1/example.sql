.mode json
.mode column
.mode table

.once out.txt
.tables
.schema

select count(*) from (select "Part Number","Description" from RES union select "Part Number","Description" from CAP);

select count(*) from table1 where exists (select * from table2 where table1.SAP == table2.SAP);
select count(*) from table where SAP in (select SAP from table2);

select count(*) from table1 where exists (select * from table2 where table1.SAP == table2.SAP);
select count(t1.SAP) from table as t1 where t1.SAP in (select t2.SAP from table2 as t2);

select count(*) from table1 where exists (select NULL from table2 where table1.SAP = table2.SAP);

