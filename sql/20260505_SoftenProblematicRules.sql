-- BaziDirectRules 軟化處理（2026-05-05）
-- 移除對客人冒犯性語言：屠戶獄官/酒店/玩物/膽小懦弱/克兄弟/手足刑傷/女性放蕩

UPDATE "BaziDirectRules" SET "Content" = '月柱傷官旺，月令為傷官，主兄弟姐妹緣份較淡，各自獨立發展為宜。' WHERE "Id" = 101;
UPDATE "BaziDirectRules" SET "Content" = '月柱干支為官殺的，月柱為兄弟宮，兄弟姐妹間宜各自獨立，比劫旺時有兄弟姐妹有當官、管事或富貴的一面。' WHERE "Id" = 103;
UPDATE "BaziDirectRules" SET "Content" = '戊寅、己卯日生人，與兄弟姐妹緣份較淡，宜各自獨立發展。' WHERE "Id" = 117;
UPDATE "BaziDirectRules" SET "Content" = '月支為殺，自己定是長子，或獨子，手足之間各自發展為宜。' WHERE "Id" = 121;
UPDATE "BaziDirectRules" SET "Content" = '男命財為喜用且在命局中佔很多分量，命中不見官星或官星受制，尚須傷官透出或居月而無受制，感情方面需多加謹慎，宜專一用情。' WHERE "Id" = 169;
UPDATE "BaziDirectRules" SET "Content" = '女命官殺強旺，婚姻需多用心經營，宜選擇性格穩重的伴侶，避免感情波折。' WHERE "Id" = 170;
UPDATE "BaziDirectRules" SET "Content" = '四柱地支全是辰戌，與公務、法律、農業相關行業較有緣份。' WHERE "Id" = 176;
UPDATE "BaziDirectRules" SET "Content" = '柱中有辛、丁、巳，支有酉、亥、未者，宜從事餐飲服務業或流通零售業。' WHERE "Id" = 188;
UPDATE "BaziDirectRules" SET "Content" = '官星為用神，八字傷官旺，逢官星透出流年，宜謹言慎行，避免法律糾紛或文件爭議。' WHERE "Id" = 235;
UPDATE "BaziDirectRules" SET "Content" = '八字官殺太多：外界壓力偏大，行事宜低調謹慎，避免樹大招風；身體較易疲勞，注意定期保養；女命婚姻需多溝通，宜選性情平和的伴侶。' WHERE "Id" = 237;
UPDATE "BaziDirectRules" SET "Content" = '八字食傷太多：言語表達能力強，但需注意言多必失，避免說話傷人；個性自我，宜學習傾聽；感情方面需專一穩重；不適合受制度約束的環境，宜自由發揮創意的工作。' WHERE "Id" = 240;

-- ziwei_patterns_144 軟化處理
UPDATE "ziwei_patterns_144"
SET "content" = replace("content", '所以容易被人家包養。', '感情上宜謹慎選擇，避免依賴關係。')
WHERE "content" ILIKE '%包養%';

UPDATE "ziwei_patterns_144"
SET "content" = replace("content",
  '此種命格男人宜從事公關、演藝人員、酒店、牛郎、或經營與女人有關之行業。女人則走演藝界、公關、服飾、酒店、餐飲為佳。',
  '此種命格適合從事公關、演藝、服飾、餐飲或創意相關行業，人際魅力為最大資產。')
WHERE "content" ILIKE '%牛郎%';

UPDATE "ziwei_patterns_144"
SET "content" = replace("content", '到酒店喝酒談生意。', '喜藉應酬場合拓展人脈、談生意。')
WHERE "content" ILIKE '%到酒店喝酒談生意%';
