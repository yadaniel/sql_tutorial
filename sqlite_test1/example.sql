.mode json
select count(*) from (select "Part Number","Description" from RES union select "Part Number","Description" from CAP);

