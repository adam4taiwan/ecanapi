-- Fix: 亡神 LookupMap 巳酉丑/亥卯未 值對調 bug
-- 正確：巳酉丑→申, 亥卯未→寅（原本寫反）

UPDATE "BaziJingShenSha"
SET "LookupMap" = '{"申子辰":"亥","寅午戌":"巳","巳酉丑":"申","亥卯未":"寅"}'
WHERE "Name" = '亡神';
