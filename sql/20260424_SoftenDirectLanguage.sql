-- 20260424 - 軟化命書直接用詞（死/夭折/早亡）+ 擎羊陀羅相夾田宅宮文字修正

-- === BaziDirectRules - ParentInfo ===
UPDATE "BaziDirectRules" SET "Content" = '年干為比劫不利父，年支為財星不利母，若幼年行比劫運、身旺，主父亡先行。'
WHERE "Condition" = '年干比劫財星論父母';

UPDATE "BaziDirectRules" SET "Content" = '四柱純陽，印衰，母親緣薄先行。時干克年干者少年別母，時干與年干不克者中年別母。'
WHERE "Condition" = '四柱純陽印衰母早喪';

UPDATE "BaziDirectRules" SET "Content" = '四柱純陰，財衰，父親緣薄先行。時干克年干者少年別父，時干與年干不克者中年別父。'
WHERE "Condition" = '四柱純陰財衰父早喪';

UPDATE "BaziDirectRules" SET "Content" = '年、月、日、時、胎支皆克干者，為父母緣薄難守之命。'
WHERE "Condition" = '年月日時胎支皆克干父母早喪';

UPDATE "BaziDirectRules" SET "Content" = '四柱中有三柱納音克胎納音者，主父母早年別離。'
WHERE "Condition" = '三柱納音克胎納音父母雙亡';

-- === BaziDirectRules - SiblingInfo ===
UPDATE "BaziDirectRules" SET "Content" = '月柱傷官旺，月令為傷官，為上不招下不招，主此人克兄弟姐妹。'
WHERE "Condition" = '月柱傷官旺克兄弟';

UPDATE "BaziDirectRules" SET "Content" = '地支本氣為官殺而中余氣藏有比劫的，十有八九兄弟姐妹中易有變故現象。'
WHERE "Condition" = '地支本氣官殺中余氣比劫兄弟夭亡';

-- === BaziDirectRules - ChildInfo ===
UPDATE "BaziDirectRules" SET "Content" = '女命，時柱坐梟、印必克子女，多流產、難產，子女緣薄。'
WHERE "Condition" = '女命時柱坐梟印克子女';

UPDATE "BaziDirectRules" SET "Content" = '時上坐梟，年月透財，女人有子亦難周全。'
WHERE "Condition" = '時上坐梟年月透財女人有子不死也傷';

UPDATE "BaziDirectRules" SET "Content" = '日、時相沖，中晚年子女緣份有損。'
WHERE "Condition" = '日時相沖中晚年喪子之憂';

-- === ziwei_patterns_144 - 田宅宮擎羊陀羅相夾文字修正 ===
UPDATE public.ziwei_patterns_144
SET content = REPLACE(content, '祖業難留承，破財賣田宅。', '祖業難承，有產化財用。')
WHERE content LIKE '%祖業難留承%';
