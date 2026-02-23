-- ============================================================
-- Song Library & Selection History Import
-- Generated: 2026-02-23T03:07:00.856Z
-- Source: \\ds1821\Backups\maple ridge\2023 backlog
-- Folders scanned: 12
-- Unique songs: 38
-- Total selections: 38
-- ============================================================


-- ============================================================
-- SONGS
-- ============================================================

-- Becky Sipla Music
-- -----------------
-- 149 - Why Do I Sing About Jesus
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Why Do I Sing About Jesus', 'why do i sing about jesus', '149', 'Becky Sipla Music' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- He Looked Beyond my Faults
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'He Looked Beyond my Faults', 'he looked beyond my faults', NULL, 'Becky Sipla Music' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- His Eye Is on the Sparrow
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'His Eye Is on the Sparrow', 'his eye is on the sparrow', NULL, 'Becky Sipla Music' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- I Am Resolved
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'I Am Resolved', 'i am resolved', NULL, 'Becky Sipla Music' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- O Mighty God, When I Behold the Wonder
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'O Mighty God, When I Behold the Wonder', 'o mighty god when i behold the wonder', NULL, 'Becky Sipla Music' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- Saved by the Blood
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Saved by the Blood', 'saved by the blood', NULL, 'Becky Sipla Music' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- When He Cometh
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'When He Cometh', 'when he cometh', NULL, 'Becky Sipla Music' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- Wonderful World
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Wonderful World', 'wonderful world', NULL, 'Becky Sipla Music' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );

-- Christmas
-- ---------
-- Beautiful Star of Bethlehem
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Beautiful Star of Bethlehem', 'beautiful star of bethlehem', NULL, 'Christmas' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- The Birthday of a King
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'The Birthday of a King', 'the birthday of a king', NULL, 'Christmas' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );

-- Hymnal
-- ------
-- 013 - We Gather Together
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'We Gather Together', 'we gather together', '013', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 031 - Love Devine
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Love Devine', 'love devine', '031', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 033 - Bless The Lord
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Bless The Lord', 'bless the lord', '033', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 046 - Praise Ye The Triune God
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Praise Ye The Triune God', 'praise ye the triune god', '046', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 047 - Glory Be to God on High
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Glory Be to God on High', 'glory be to god on high', '047', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 048 - Come Thou All Mighty King
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Come Thou All Mighty King', 'come thou all mighty king', '048', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 049 - Holy God We Praise Thy Name
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Holy God We Praise Thy Name', 'holy god we praise thy name', '049', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 055 - Immortal, Invisible
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Immortal, Invisible', 'immortal invisible', '055', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 056 - This Is My Fathers World
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'This Is My Fathers World', 'this is my fathers world', '056', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 110 - Thanks to God
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Thanks to God', 'thanks to god', '110', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 111 - Come, Ye Thankful People
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Come, Ye Thankful People', 'come ye thankful people', '111', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 168 - Come and Praise
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Come and Praise', 'come and praise', '168', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 170 - Joy to the World
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Joy to the World', 'joy to the world', '170', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 171 - O Come O Come Emmanuel
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'O Come O Come Emmanuel', 'o come o come emmanuel', '171', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 174 - Lift Up Your Heads
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Lift Up Your Heads', 'lift up your heads', '174', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 191 - While By Our Sheep
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'While By Our Sheep', 'while by our sheep', '191', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 193 - It Came upon the Midnight Clear
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'It Came upon the Midnight Clear', 'it came upon the midnight clear', '193', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 198 - Good Christian Men Rejoice
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Good Christian Men Rejoice', 'good christian men rejoice', '198', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 329 - Anywhere With Jesus
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Anywhere With Jesus', 'anywhere with jesus', '329', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 430 - Count Your Blessings
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Count Your Blessings', 'count your blessings', '430', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 536 - My Tribute
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'My Tribute', 'my tribute', '536', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 541 - In The Garden
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'In The Garden', 'in the garden', '541', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 542 - Jesus Loves Even Me
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Jesus Loves Even Me', 'jesus loves even me', '542', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 546 - I Will Sing The Wondrous Story
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'I Will Sing The Wondrous Story', 'i will sing the wondrous story', '546', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 548 - Glory To His Name
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Glory To His Name', 'glory to his name', '548', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 549 - He's Everything to Me
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'He''s Everything to Me', 'hes everything to me', '549', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 550 - He Lives
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'He Lives', 'he lives', '550', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 558 - Springs of Living Water
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Springs of Living Water', 'springs of living water', '558', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );


-- ============================================================
-- SONG SELECTIONS
-- ============================================================

-- Aug 6, 2023
--   Song 1: 031 - Love Devine
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230806', 'song', 1 FROM songs WHERE normalized_name = 'love devine' AND number = '031' AND book = 'Hymnal' LIMIT 1;
--   Song 2: 541 - In The Garden
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230806', 'song', 2 FROM songs WHERE normalized_name = 'in the garden' AND number = '541' AND book = 'Hymnal' LIMIT 1;
--   Song 3: When He Cometh
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230806', 'song', 3 FROM songs WHERE normalized_name = 'when he cometh' AND book = 'Becky Sipla Music' LIMIT 1;

-- Aug 13, 2023
--   Song 1: 033 - Bless The Lord
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230813', 'song', 1 FROM songs WHERE normalized_name = 'bless the lord' AND number = '033' AND book = 'Hymnal' LIMIT 1;
--   Song 2: 542 - Jesus Loves Even Me
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230813', 'song', 2 FROM songs WHERE normalized_name = 'jesus loves even me' AND number = '542' AND book = 'Hymnal' LIMIT 1;
--   Song 3: Wonderful World
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230813', 'song', 3 FROM songs WHERE normalized_name = 'wonderful world' AND book = 'Becky Sipla Music' LIMIT 1;

-- Sep 10, 2023
--   Song 1: 046 - Praise Ye The Triune God
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230910', 'song', 1 FROM songs WHERE normalized_name = 'praise ye the triune god' AND number = '046' AND book = 'Hymnal' LIMIT 1;
--   Song 2: 546 - I Will Sing The Wondrous Story
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230910', 'song', 2 FROM songs WHERE normalized_name = 'i will sing the wondrous story' AND number = '546' AND book = 'Hymnal' LIMIT 1;
--   Song 3: O Mighty God, When I Behold the Wonder
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230910', 'song', 3 FROM songs WHERE normalized_name = 'o mighty god when i behold the wonder' AND book = 'Becky Sipla Music' LIMIT 1;

-- Sep 17, 2023
--   Song 1: 047 - Glory Be to God on High
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230917', 'song', 1 FROM songs WHERE normalized_name = 'glory be to god on high' AND number = '047' AND book = 'Hymnal' LIMIT 1;
--   Song 2: 329 - Anywhere With Jesus
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230917', 'song', 2 FROM songs WHERE normalized_name = 'anywhere with jesus' AND number = '329' AND book = 'Hymnal' LIMIT 1;
--   Song 3: 548 - Glory To His Name
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230917', 'song', 3 FROM songs WHERE normalized_name = 'glory to his name' AND number = '548' AND book = 'Hymnal' LIMIT 1;
--   Song 4: Saved by the Blood
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230917', 'song', 4 FROM songs WHERE normalized_name = 'saved by the blood' AND book = 'Becky Sipla Music' LIMIT 1;

-- Sep 24, 2023
--   Song 1: 048 - Come Thou All Mighty King
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230924', 'song', 1 FROM songs WHERE normalized_name = 'come thou all mighty king' AND number = '048' AND book = 'Hymnal' LIMIT 1;
--   Song 2: 549 - He's Everything to Me
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230924', 'song', 2 FROM songs WHERE normalized_name = 'hes everything to me' AND number = '549' AND book = 'Hymnal' LIMIT 1;
--   Song 3: I Am Resolved
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230924', 'song', 3 FROM songs WHERE normalized_name = 'i am resolved' AND book = 'Becky Sipla Music' LIMIT 1;

-- Oct 1, 2023
--   Song 1: 049 - Holy God We Praise Thy Name
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20231001', 'song', 1 FROM songs WHERE normalized_name = 'holy god we praise thy name' AND number = '049' AND book = 'Hymnal' LIMIT 1;
--   Song 2: 550 - He Lives
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20231001', 'song', 2 FROM songs WHERE normalized_name = 'he lives' AND number = '550' AND book = 'Hymnal' LIMIT 1;
--   Song 3: He Looked Beyond my Faults
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20231001', 'song', 3 FROM songs WHERE normalized_name = 'he looked beyond my faults' AND book = 'Becky Sipla Music' LIMIT 1;
--   Song 4: His Eye Is on the Sparrow
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20231001', 'song', 4 FROM songs WHERE normalized_name = 'his eye is on the sparrow' AND book = 'Becky Sipla Music' LIMIT 1;

-- Nov 12, 2023
--   Song 1: 055 - Immortal, Invisible
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20231112', 'song', 1 FROM songs WHERE normalized_name = 'immortal invisible' AND number = '055' AND book = 'Hymnal' LIMIT 1;
--   Song 2: 149 - Why Do I Sing About Jesus
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20231112', 'song', 2 FROM songs WHERE normalized_name = 'why do i sing about jesus' AND number = '149' AND book = 'Becky Sipla Music' LIMIT 1;
--   Song 3: 558 - Springs of Living Water
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20231112', 'song', 3 FROM songs WHERE normalized_name = 'springs of living water' AND number = '558' AND book = 'Hymnal' LIMIT 1;

-- Nov 19, 2023
--   Song 1: 013 - We Gather Together
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20231119', 'song', 1 FROM songs WHERE normalized_name = 'we gather together' AND number = '013' AND book = 'Hymnal' LIMIT 1;
--   Song 2: 056 - This Is My Fathers World
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20231119', 'song', 2 FROM songs WHERE normalized_name = 'this is my fathers world' AND number = '056' AND book = 'Hymnal' LIMIT 1;
--   Song 3: 110 - Thanks to God
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20231119', 'song', 3 FROM songs WHERE normalized_name = 'thanks to god' AND number = '110' AND book = 'Hymnal' LIMIT 1;

-- Nov 26, 2023
--   Song 1: 111 - Come, Ye Thankful People
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20231126', 'song', 1 FROM songs WHERE normalized_name = 'come ye thankful people' AND number = '111' AND book = 'Hymnal' LIMIT 1;
--   Song 2: 430 - Count Your Blessings
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20231126', 'song', 2 FROM songs WHERE normalized_name = 'count your blessings' AND number = '430' AND book = 'Hymnal' LIMIT 1;
--   Song 3: 536 - My Tribute
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20231126', 'song', 3 FROM songs WHERE normalized_name = 'my tribute' AND number = '536' AND book = 'Hymnal' LIMIT 1;

-- Dec 3, 2023
--   Song 1: 168 - Come and Praise
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20231203', 'song', 1 FROM songs WHERE normalized_name = 'come and praise' AND number = '168' AND book = 'Hymnal' LIMIT 1;
--   Song 2: 174 - Lift Up Your Heads
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20231203', 'song', 2 FROM songs WHERE normalized_name = 'lift up your heads' AND number = '174' AND book = 'Hymnal' LIMIT 1;
--   Song 3: 198 - Good Christian Men Rejoice
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20231203', 'song', 3 FROM songs WHERE normalized_name = 'good christian men rejoice' AND number = '198' AND book = 'Hymnal' LIMIT 1;

-- Dec 17, 2023
--   Song 1: 170 - Joy to the World
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20231217', 'song', 1 FROM songs WHERE normalized_name = 'joy to the world' AND number = '170' AND book = 'Hymnal' LIMIT 1;
--   Song 2: 193 - It Came upon the Midnight Clear
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20231217', 'song', 2 FROM songs WHERE normalized_name = 'it came upon the midnight clear' AND number = '193' AND book = 'Hymnal' LIMIT 1;
--   Song 3: Beautiful Star of Bethlehem
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20231217', 'song', 3 FROM songs WHERE normalized_name = 'beautiful star of bethlehem' AND book = 'Christmas' LIMIT 1;

-- Dec 24, 2023
--   Song 1: 171 - O Come O Come Emmanuel
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20231224', 'song', 1 FROM songs WHERE normalized_name = 'o come o come emmanuel' AND number = '171' AND book = 'Hymnal' LIMIT 1;
--   Song 2: 191 - While By Our Sheep
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20231224', 'song', 2 FROM songs WHERE normalized_name = 'while by our sheep' AND number = '191' AND book = 'Hymnal' LIMIT 1;
--   Song 3: The Birthday of a King
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20231224', 'song', 3 FROM songs WHERE normalized_name = 'the birthday of a king' AND book = 'Christmas' LIMIT 1;
