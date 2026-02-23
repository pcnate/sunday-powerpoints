-- ============================================================
-- Song Library & Selection History Import
-- Generated: 2026-02-23T03:04:55.560Z
-- Source: \\ds1821\Backups\maple ridge\2023
-- Folders scanned: 44
-- Unique songs: 127
-- Total selections: 136
-- ============================================================


-- ============================================================
-- SONGS
-- ============================================================

-- Becky Sipla Music
-- -----------------
-- 038 - Nothing is Impossible
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Nothing is Impossible', 'nothing is impossible', '038', 'Becky Sipla Music' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 042 - Take Time to Be Holy
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Take Time to Be Holy', 'take time to be holy', '042', 'Becky Sipla Music' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 108 - All Glory, Laud and Honor
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'All Glory, Laud and Honor', 'all glory laud and honor', '108', 'Becky Sipla Music' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 115 - Ivory Palaces
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Ivory Palaces', 'ivory palaces', '115', 'Becky Sipla Music' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 121 - Blessed Calvary
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Blessed Calvary', 'blessed calvary', '121', 'Becky Sipla Music' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 149 - Why Do I Sing About Jesus
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Why Do I Sing About Jesus', 'why do i sing about jesus', '149', 'Becky Sipla Music' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 154 - A New Name In Glory
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'A New Name In Glory', 'a new name in glory', '154', 'Becky Sipla Music' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 157 - What a Gathering
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'What a Gathering', 'what a gathering', '157', 'Becky Sipla Music' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 199 - Christ Receiveth Sinful Men
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Christ Receiveth Sinful Men', 'christ receiveth sinful men', '199', 'Becky Sipla Music' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 205 - He Included Me
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'He Included Me', 'he included me', '205', 'Becky Sipla Music' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 210 - Saved by the Blood
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Saved by the Blood', 'saved by the blood', '210', 'Becky Sipla Music' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 229 - Our God Reigns
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Our God Reigns', 'our god reigns', '229', 'Becky Sipla Music' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 322 - The Bible Stands
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'The Bible Stands', 'the bible stands', '322', 'Becky Sipla Music' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 381 - Is Your All on the Altar
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Is Your All on the Altar', 'is your all on the altar', '381', 'Becky Sipla Music' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 426 - Throw Out the Life-Line
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Throw Out the Life-Line', 'throw out the life line', '426', 'Becky Sipla Music' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 436 - Bring Them In
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Bring Them In', 'bring them in', '436', 'Becky Sipla Music' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 448 - He Ransomed Me
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'He Ransomed Me', 'he ransomed me', '448', 'Becky Sipla Music' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 454 - Glorious is Thy Name Most Holy
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Glorious is Thy Name Most Holy', 'glorious is thy name most holy', '454', 'Becky Sipla Music' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 464 - I Will Praise Him
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'I Will Praise Him', 'i will praise him', '464', 'Becky Sipla Music' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 472 - Heavenly Sunlight
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Heavenly Sunlight', 'heavenly sunlight', '472', 'Becky Sipla Music' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 476 - It Is Glory Just to Walk with Him
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'It Is Glory Just to Walk with Him', 'it is glory just to walk with him', '476', 'Becky Sipla Music' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- Alleluia! Sing to Jesus
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Alleluia! Sing to Jesus', 'alleluia sing to jesus', NULL, 'Becky Sipla Music' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- And Can It Be
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'And Can It Be', 'and can it be', NULL, 'Becky Sipla Music' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- Be Thou Exalted
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Be Thou Exalted', 'be thou exalted', NULL, 'Becky Sipla Music' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- Bringing in the Sheaves
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Bringing in the Sheaves', 'bringing in the sheaves', NULL, 'Becky Sipla Music' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- Faith of our Fathers
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Faith of our Fathers', 'faith of our fathers', NULL, 'Becky Sipla Music' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- Faith of Our Mothers
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Faith of Our Mothers', 'faith of our mothers', NULL, 'Becky Sipla Music' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- Footsteps of Jesus
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Footsteps of Jesus', 'footsteps of jesus', NULL, 'Becky Sipla Music' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- Have I Done My Best for Jesus
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Have I Done My Best for Jesus', 'have i done my best for jesus', NULL, 'Becky Sipla Music' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- He is Able to Deliver Thee
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'He is Able to Deliver Thee', 'he is able to deliver thee', NULL, 'Becky Sipla Music' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- How Great is Our God
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'How Great is Our God', 'how great is our god', NULL, 'Becky Sipla Music' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- Its Just Like His Great Love
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Its Just Like His Great Love', 'its just like his great love', NULL, 'Becky Sipla Music' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- Lavish Love, Abundant Beauty
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Lavish Love, Abundant Beauty', 'lavish love abundant beauty', NULL, 'Becky Sipla Music' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- Let Jesus Come Into Your Heart
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Let Jesus Come Into Your Heart', 'let jesus come into your heart', NULL, 'Becky Sipla Music' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- Let the Lower Lights be Burning
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Let the Lower Lights be Burning', 'let the lower lights be burning', NULL, 'Becky Sipla Music' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- Lift Up, Lift Up Your Voices Now
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Lift Up, Lift Up Your Voices Now', 'lift up lift up your voices now', NULL, 'Becky Sipla Music' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- Prince of Peace
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Prince of Peace', 'prince of peace', NULL, 'Becky Sipla Music' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- The Beauty of Jesus
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'The Beauty of Jesus', 'the beauty of jesus', NULL, 'Becky Sipla Music' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- The Bible Stands
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'The Bible Stands', 'the bible stands', NULL, 'Becky Sipla Music' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- The Lily of the Valley
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'The Lily of the Valley', 'the lily of the valley', NULL, 'Becky Sipla Music' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- The Name of Jesus
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'The Name of Jesus', 'the name of jesus', NULL, 'Becky Sipla Music' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- The Star-Spangled Banner
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'The Star-Spangled Banner', 'the star spangled banner', NULL, 'Becky Sipla Music' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- We Have an Anchor
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'We Have an Anchor', 'we have an anchor', NULL, 'Becky Sipla Music' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- What a Gathering
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'What a Gathering', 'what a gathering', NULL, 'Becky Sipla Music' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- Yesterday, Today, Forever
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Yesterday, Today, Forever', 'yesterday today forever', NULL, 'Becky Sipla Music' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );

-- Choruses
-- --------
-- Prince Of Peace
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Prince Of Peace', 'prince of peace', NULL, 'Choruses' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );

-- Hymnal
-- ------
-- 035 - Come Thou Fount
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Come Thou Fount', 'come thou fount', '035', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 037 - All Creatures of Our God and King
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'All Creatures of Our God and King', 'all creatures of our god and king', '037', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 041 - Holy, Holy, Holy
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Holy, Holy, Holy', 'holy holy holy', '041', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 051 - Holy Holy
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Holy Holy', 'holy holy', '051', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 053 - Joyful Joyful We Adore Thee
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Joyful Joyful We Adore Thee', 'joyful joyful we adore thee', '053', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 054 - Great is Thy Faithfulness
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Great is Thy Faithfulness', 'great is thy faithfulness', '054', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 055 - Immortal, Invisible
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Immortal, Invisible', 'immortal invisible', '055', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 056 - This Is My Fathers World
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'This Is My Fathers World', 'this is my fathers world', '056', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 057 - Our Great Savior
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Our Great Savior', 'our great savior', '057', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 058 - Fairest Lord Jesus
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Fairest Lord Jesus', 'fairest lord jesus', '058', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 060 - May Jesus Christ Be Praised
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'May Jesus Christ Be Praised', 'may jesus christ be praised', '060', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 062 - O for a Thousand Tongues
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'O for a Thousand Tongues', 'o for a thousand tongues', '062', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 063 - Praise Him! Praise  Him!
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Praise Him! Praise  Him!', 'praise him praise him', '063', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 064 - Come, Christans, Join to Sing
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Come, Christans, Join to Sing', 'come christans join to sing', '064', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 066 - There's Something About That Name
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'There''s Something About That Name', 'theres something about that name', '066', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 067 - O Come Let Us Adore Him
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'O Come Let Us Adore Him', 'o come let us adore him', '067', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 069 - Jesus We Just Want to Thank You
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Jesus We Just Want to Thank You', 'jesus we just want to thank you', '069', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 070 - I Am His and He Is Mine
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'I Am His and He Is Mine', 'i am his and he is mine', '070', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 071 - Join All the Glorious Names
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Join All the Glorious Names', 'join all the glorious names', '071', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 072 - Thou Art Worthy
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Thou Art Worthy', 'thou art worthy', '072', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 092 - Channels Only
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Channels Only', 'channels only', '092', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 114 - Another Year is Dawning
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Another Year is Dawning', 'another year is dawning', '114', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 167 - One Day
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'One Day', 'one day', '167', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 169 - Thou Didst Leave Thy Throne
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Thou Didst Leave Thy Throne', 'thou didst leave thy throne', '169', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 172 - Come, Thou Long-Expected Jesus
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Come, Thou Long-Expected Jesus', 'come thou long expected jesus', '172', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 179 - Angels We Have Heard On High
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Angels We Have Heard On High', 'angels we have heard on high', '179', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 182 - Angels from the Realms of Glory
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Angels from the Realms of Glory', 'angels from the realms of glory', '182', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 194 - As With Gladness Men Of Old
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'As With Gladness Men Of Old', 'as with gladness men of old', '194', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 217 - Why
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Why', 'why', '217', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 224 - When I Survey The Wondrous Cross
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'When I Survey The Wondrous Cross', 'when i survey the wondrous cross', '224', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 227 - At the Cross
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'At the Cross', 'at the cross', '227', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 231 - Christ the Lord Is Risen Today
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Christ the Lord Is Risen Today', 'christ the lord is risen today', '231', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 243 - Crown Him with Many Crowns
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Crown Him with Many Crowns', 'crown him with many crowns', '243', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 252 - Christ Returneth
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Christ Returneth', 'christ returneth', '252', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 267 - Burdens Are Lifted At Calvary
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Burdens Are Lifted At Calvary', 'burdens are lifted at calvary', '267', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 301 - Softly and Tenderly
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Softly and Tenderly', 'softly and tenderly', '301', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 321 - It Is Well With My Soul
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'It Is Well With My Soul', 'it is well with my soul', '321', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 440 - Follow On
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Follow On', 'follow on', '440', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 486 - My Country Tis of Thee
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'My Country Tis of Thee', 'my country tis of thee', '486', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 497 - O Zion Haste
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'O Zion Haste', 'o zion haste', '497', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 498 - Christ for the World We Sing
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Christ for the World We Sing', 'christ for the world we sing', '498', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 501 - So I Send You
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'So I Send You', 'so i send you', '501', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 502 - We've a Story to Tell
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'We''ve a Story to Tell', 'weve a story to tell', '502', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 505 - Send The Light
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Send The Light', 'send the light', '505', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 507 - Pass It On
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Pass It On', 'pass it on', '507', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 513 - I'll Go Where You Want Me To Go
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'I''ll Go Where You Want Me To Go', 'ill go where you want me to go', '513', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 515 - Rescue the Perishing
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Rescue the Perishing', 'rescue the perishing', '515', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 517 - Freely, Freely
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Freely, Freely', 'freely freely', '517', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 518 - It Took A Miracle
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'It Took A Miracle', 'it took a miracle', '518', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 519 - Love Lifted Me
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Love Lifted Me', 'love lifted me', '519', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 520 - O How I Love Jesus
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'O How I Love Jesus', 'o how i love jesus', '520', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 521 - Ive Found a Friend
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Ive Found a Friend', 'ive found a friend', '521', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 522 - My Savior's Love
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'My Savior''s Love', 'my saviors love', '522', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 523 - Saved Saved
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Saved Saved', 'saved saved', '523', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 524 - At Calvary
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'At Calvary', 'at calvary', '524', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 525 - Heaven Came Down
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Heaven Came Down', 'heaven came down', '525', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 526 - Victory in Jesus
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Victory in Jesus', 'victory in jesus', '526', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 528 - Without Him
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Without Him', 'without him', '528', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 529 - No Not One
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'No Not One', 'no not one', '529', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 531 - What A Wonderful Savior
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'What A Wonderful Savior', 'what a wonderful savior', '531', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 532 - He Touched Me
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'He Touched Me', 'he touched me', '532', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 533 - I Believe in Miracles
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'I Believe in Miracles', 'i believe in miracles', '533', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 534 - Love Found A Way
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Love Found A Way', 'love found a way', '534', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 535 - Satisfied
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Satisfied', 'satisfied', '535', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 537 - New Life
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'New Life', 'new life', '537', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 538 - Now I Belong to Jesus
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Now I Belong to Jesus', 'now i belong to jesus', '538', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 539 - I Will Sing of My Redeemer
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'I Will Sing of My Redeemer', 'i will sing of my redeemer', '539', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 543 - Through It All
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Through It All', 'through it all', '543', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 545 - He Lifted Me
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'He Lifted Me', 'he lifted me', '545', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 550 - He Lives
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'He Lives', 'he lives', '550', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 552 - In My Heart There Rings a Melody
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'In My Heart There Rings a Melody', 'in my heart there rings a melody', '552', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 553 - Isnt The Love Of Jesus Something Wonderful
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Isnt The Love Of Jesus Something Wonderful', 'isnt the love of jesus something wonderful', '553', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 554 - The Old Rugged Cross
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'The Old Rugged Cross', 'the old rugged cross', '554', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 555 - Jesus Loves Me
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Jesus Loves Me', 'jesus loves me', '555', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 557 - Redeemed
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Redeemed', 'redeemed', '557', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 558 - Springs of Living Water
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Springs of Living Water', 'springs of living water', '558', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 559 - Sunshine in My Soul
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Sunshine in My Soul', 'sunshine in my soul', '559', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 561 - Since Jesus Came into My Heart
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'Since Jesus Came into My Heart', 'since jesus came into my heart', '561', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- 563 - All That Thrills My Soul
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'All That Thrills My Soul', 'all that thrills my soul', '563', 'Hymnal' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );

-- Pictures
-- --------
-- 07 The Women I Come From [Live].mp3
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( '07 The Women I Come From [Live].mp3', 'the women i come from livemp3', NULL, 'Pictures' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );
-- The Women I Come From
INSERT INTO songs ( name, normalized_name, number, book ) VALUES ( 'The Women I Come From', 'the women i come from', NULL, 'Pictures' ) ON DUPLICATE KEY UPDATE number = COALESCE( songs.number, VALUES( number ) ), book = COALESCE( songs.book, VALUES( book ) );


-- ============================================================
-- SONG SELECTIONS
-- ============================================================

-- Jan 1, 2023
--   Song 1: 115 - Ivory Palaces
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230101', 'song', 1 FROM songs WHERE normalized_name = 'ivory palaces' AND number = '115' AND book = 'Becky Sipla Music' LIMIT 1;
--   Song 2: 167 - One Day
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230101', 'song', 2 FROM songs WHERE normalized_name = 'one day' AND number = '167' AND book = 'Hymnal' LIMIT 1;
--   Song 3: 210 - Saved by the Blood
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230101', 'song', 3 FROM songs WHERE normalized_name = 'saved by the blood' AND number = '210' AND book = 'Becky Sipla Music' LIMIT 1;

-- Jan 8, 2023
--   Song 1: 038 - Nothing is Impossible
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230108', 'song', 1 FROM songs WHERE normalized_name = 'nothing is impossible' AND number = '038' AND book = 'Becky Sipla Music' LIMIT 1;
--   Song 2: 381 - Is Your All on the Altar
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230108', 'song', 2 FROM songs WHERE normalized_name = 'is your all on the altar' AND number = '381' AND book = 'Becky Sipla Music' LIMIT 1;
--   Song 3: 454 - Glorious is Thy Name Most Holy
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230108', 'song', 3 FROM songs WHERE normalized_name = 'glorious is thy name most holy' AND number = '454' AND book = 'Becky Sipla Music' LIMIT 1;

-- Jan 15, 2023
--   Song 1: 055 - Immortal, Invisible
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230115', 'song', 1 FROM songs WHERE normalized_name = 'immortal invisible' AND number = '055' AND book = 'Hymnal' LIMIT 1;
--   Song 2: 229 - Our God Reigns
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230115', 'song', 2 FROM songs WHERE normalized_name = 'our god reigns' AND number = '229' AND book = 'Becky Sipla Music' LIMIT 1;
--   Song 3: 464 - I Will Praise Him
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230115', 'song', 3 FROM songs WHERE normalized_name = 'i will praise him' AND number = '464' AND book = 'Becky Sipla Music' LIMIT 1;

-- Jan 22, 2023
--   Song 1: 056 - This Is My Fathers World
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230122', 'song', 1 FROM songs WHERE normalized_name = 'this is my fathers world' AND number = '056' AND book = 'Hymnal' LIMIT 1;
--   Song 2: 157 - What a Gathering
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230122', 'song', 2 FROM songs WHERE normalized_name = 'what a gathering' AND number = '157' AND book = 'Becky Sipla Music' LIMIT 1;
--   Song 3: 497 - O Zion Haste
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230122', 'song', 3 FROM songs WHERE normalized_name = 'o zion haste' AND number = '497' AND book = 'Hymnal' LIMIT 1;

-- Jan 29, 2023
--   Song 1: 057 - Our Great Savior
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230129', 'song', 1 FROM songs WHERE normalized_name = 'our great savior' AND number = '057' AND book = 'Hymnal' LIMIT 1;
--   Song 2: 149 - Why Do I Sing About Jesus
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230129', 'song', 2 FROM songs WHERE normalized_name = 'why do i sing about jesus' AND number = '149' AND book = 'Becky Sipla Music' LIMIT 1;
--   Song 3: 267 - Burdens Are Lifted At Calvary
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230129', 'song', 3 FROM songs WHERE normalized_name = 'burdens are lifted at calvary' AND number = '267' AND book = 'Hymnal' LIMIT 1;

-- Feb 5, 2023
--   Song 1: 058 - Fairest Lord Jesus
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230205', 'song', 1 FROM songs WHERE normalized_name = 'fairest lord jesus' AND number = '058' AND book = 'Hymnal' LIMIT 1;
--   Song 2: 322 - The Bible Stands
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230205', 'song', 2 FROM songs WHERE normalized_name = 'the bible stands' AND number = '322' AND book = 'Becky Sipla Music' LIMIT 1;
--   Song 3: 498 - Christ for the World We Sing
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230205', 'song', 3 FROM songs WHERE normalized_name = 'christ for the world we sing' AND number = '498' AND book = 'Hymnal' LIMIT 1;

-- Feb 12, 2023
--   Song 1: 035 - Come Thou Fount
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230212', 'song', 1 FROM songs WHERE normalized_name = 'come thou fount' AND number = '035' AND book = 'Hymnal' LIMIT 1;
--   Song 2: 063 - Praise Him! Praise  Him!
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230212', 'song', 2 FROM songs WHERE normalized_name = 'praise him praise him' AND number = '063' AND book = 'Hymnal' LIMIT 1;
--   Song 3: 526 - Victory in Jesus
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230212', 'song', 3 FROM songs WHERE normalized_name = 'victory in jesus' AND number = '526' AND book = 'Hymnal' LIMIT 1;

-- Feb 19, 2023
--   Song 1: 060 - May Jesus Christ Be Praised
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230219', 'song', 1 FROM songs WHERE normalized_name = 'may jesus christ be praised' AND number = '060' AND book = 'Hymnal' LIMIT 1;
--   Song 2: 472 - Heavenly Sunlight
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230219', 'song', 2 FROM songs WHERE normalized_name = 'heavenly sunlight' AND number = '472' AND book = 'Becky Sipla Music' LIMIT 1;
--   Song 3: 501 - So I Send You
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230219', 'song', 3 FROM songs WHERE normalized_name = 'so i send you' AND number = '501' AND book = 'Hymnal' LIMIT 1;

-- Feb 26, 2023
--   Song 1: 448 - He Ransomed Me
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230226', 'song', 1 FROM songs WHERE normalized_name = 'he ransomed me' AND number = '448' AND book = 'Becky Sipla Music' LIMIT 1;
--   Song 2: 502 - We've a Story to Tell
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230226', 'song', 2 FROM songs WHERE normalized_name = 'weve a story to tell' AND number = '502' AND book = 'Hymnal' LIMIT 1;
--   Song 3: And Can It Be
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230226', 'song', 3 FROM songs WHERE normalized_name = 'and can it be' AND book = 'Becky Sipla Music' LIMIT 1;

-- Mar 5, 2023
--   Song 1: 042 - Take Time to Be Holy
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230305', 'song', 1 FROM songs WHERE normalized_name = 'take time to be holy' AND number = '042' AND book = 'Becky Sipla Music' LIMIT 1;
--   Song 2: 121 - Blessed Calvary
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230305', 'song', 2 FROM songs WHERE normalized_name = 'blessed calvary' AND number = '121' AND book = 'Becky Sipla Music' LIMIT 1;
--   Song 3: 199 - Christ Receiveth Sinful Men
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230305', 'song', 3 FROM songs WHERE normalized_name = 'christ receiveth sinful men' AND number = '199' AND book = 'Becky Sipla Music' LIMIT 1;

-- Mar 12, 2023
--   Song 1: 062 - O for a Thousand Tongues
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230312', 'song', 1 FROM songs WHERE normalized_name = 'o for a thousand tongues' AND number = '062' AND book = 'Hymnal' LIMIT 1;
--   Song 2: 440 - Follow On
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230312', 'song', 2 FROM songs WHERE normalized_name = 'follow on' AND number = '440' AND book = 'Hymnal' LIMIT 1;
--   Song 3: Footsteps of Jesus
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230312', 'song', 3 FROM songs WHERE normalized_name = 'footsteps of jesus' AND book = 'Becky Sipla Music' LIMIT 1;

-- Mar 19, 2023
--   Song 1: 063 - Praise Him! Praise  Him!
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230319', 'song', 1 FROM songs WHERE normalized_name = 'praise him praise him' AND number = '063' AND book = 'Hymnal' LIMIT 1;
--   Song 2: 436 - Bring Them In
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230319', 'song', 2 FROM songs WHERE normalized_name = 'bring them in' AND number = '436' AND book = 'Becky Sipla Music' LIMIT 1;
--   Song 3: 505 - Send The Light
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230319', 'song', 3 FROM songs WHERE normalized_name = 'send the light' AND number = '505' AND book = 'Hymnal' LIMIT 1;

-- Mar 26, 2023
--   Song 1: 064 - Come, Christans, Join to Sing
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230326', 'song', 1 FROM songs WHERE normalized_name = 'come christans join to sing' AND number = '064' AND book = 'Hymnal' LIMIT 1;
--   Song 2: 476 - It Is Glory Just to Walk with Him
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230326', 'song', 2 FROM songs WHERE normalized_name = 'it is glory just to walk with him' AND number = '476' AND book = 'Becky Sipla Music' LIMIT 1;
--   Song 3: 507 - Pass It On
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230326', 'song', 3 FROM songs WHERE normalized_name = 'pass it on' AND number = '507' AND book = 'Hymnal' LIMIT 1;

-- Apr 2, 2023
--   Song 1: 108 - All Glory, Laud and Honor
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230402', 'song', 1 FROM songs WHERE normalized_name = 'all glory laud and honor' AND number = '108' AND book = 'Becky Sipla Music' LIMIT 1;
--   Song 2: 513 - I'll Go Where You Want Me To Go
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230402', 'song', 2 FROM songs WHERE normalized_name = 'ill go where you want me to go' AND number = '513' AND book = 'Hymnal' LIMIT 1;
--   Song 3: The Beauty of Jesus
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230402', 'song', 3 FROM songs WHERE normalized_name = 'the beauty of jesus' AND book = 'Becky Sipla Music' LIMIT 1;

-- Apr 9, 2023
--   Song 1: 231 - Christ the Lord Is Risen Today
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230409', 'song', 1 FROM songs WHERE normalized_name = 'christ the lord is risen today' AND number = '231' AND book = 'Hymnal' LIMIT 1;
--   Song 2: 550 - He Lives
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230409', 'song', 2 FROM songs WHERE normalized_name = 'he lives' AND number = '550' AND book = 'Hymnal' LIMIT 1;
--   Song 3: How Great is Our God
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230409', 'song', 3 FROM songs WHERE normalized_name = 'how great is our god' AND book = 'Becky Sipla Music' LIMIT 1;

-- Apr 16, 2023
--   Song 1: 066 - There's Something About That Name
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230416', 'song', 1 FROM songs WHERE normalized_name = 'theres something about that name' AND number = '066' AND book = 'Hymnal' LIMIT 1;
--   Song 2: 515 - Rescue the Perishing
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230416', 'song', 2 FROM songs WHERE normalized_name = 'rescue the perishing' AND number = '515' AND book = 'Hymnal' LIMIT 1;
--   Song 3: Alleluia! Sing to Jesus
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230416', 'song', 3 FROM songs WHERE normalized_name = 'alleluia sing to jesus' AND book = 'Becky Sipla Music' LIMIT 1;

-- Apr 23, 2023
--   Song 1: 041 - Holy, Holy, Holy
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230423', 'song', 1 FROM songs WHERE normalized_name = 'holy holy holy' AND number = '041' AND book = 'Hymnal' LIMIT 1;
--   Song 2: 092 - Channels Only
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230423', 'song', 2 FROM songs WHERE normalized_name = 'channels only' AND number = '092' AND book = 'Hymnal' LIMIT 1;
--   Song 3: 321 - It Is Well With My Soul
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230423', 'song', 3 FROM songs WHERE normalized_name = 'it is well with my soul' AND number = '321' AND book = 'Hymnal' LIMIT 1;

-- Apr 30, 2023
--   Song 1: 067 - O Come Let Us Adore Him
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230430', 'song', 1 FROM songs WHERE normalized_name = 'o come let us adore him' AND number = '067' AND book = 'Hymnal' LIMIT 1;
--   Song 2: 517 - Freely, Freely
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230430', 'song', 2 FROM songs WHERE normalized_name = 'freely freely' AND number = '517' AND book = 'Hymnal' LIMIT 1;
--   Song 3: Let Jesus Come Into Your Heart
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230430', 'song', 3 FROM songs WHERE normalized_name = 'let jesus come into your heart' AND book = 'Becky Sipla Music' LIMIT 1;

-- May 7, 2023
--   Song 1: 069 - Jesus We Just Want to Thank You
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230507', 'song', 1 FROM songs WHERE normalized_name = 'jesus we just want to thank you' AND number = '069' AND book = 'Hymnal' LIMIT 1;
--   Song 2: 154 - A New Name In Glory
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230507', 'song', 2 FROM songs WHERE normalized_name = 'a new name in glory' AND number = '154' AND book = 'Becky Sipla Music' LIMIT 1;
--   Song 3: 518 - It Took A Miracle
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230507', 'song', 3 FROM songs WHERE normalized_name = 'it took a miracle' AND number = '518' AND book = 'Hymnal' LIMIT 1;

-- May 14, 2023
--   Song 1: 07 The Women I Come From [Live].mp3
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230514', 'song', 1 FROM songs WHERE normalized_name = 'the women i come from livemp3' AND book = 'Pictures' LIMIT 1;
--   Song 2: 070 - I Am His and He Is Mine
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230514', 'song', 2 FROM songs WHERE normalized_name = 'i am his and he is mine' AND number = '070' AND book = 'Hymnal' LIMIT 1;
--   Song 3: 519 - Love Lifted Me
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230514', 'song', 3 FROM songs WHERE normalized_name = 'love lifted me' AND number = '519' AND book = 'Hymnal' LIMIT 1;
--   Song 4: Faith of Our Mothers
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230514', 'song', 4 FROM songs WHERE normalized_name = 'faith of our mothers' AND book = 'Becky Sipla Music' LIMIT 1;
--   Song 5: The Women I Come From
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230514', 'song', 5 FROM songs WHERE normalized_name = 'the women i come from' AND book = 'Pictures' LIMIT 1;

-- May 21, 2023
--   Song 1: 071 - Join All the Glorious Names
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230521', 'song', 1 FROM songs WHERE normalized_name = 'join all the glorious names' AND number = '071' AND book = 'Hymnal' LIMIT 1;
--   Song 2: 243 - Crown Him with Many Crowns
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230521', 'song', 2 FROM songs WHERE normalized_name = 'crown him with many crowns' AND number = '243' AND book = 'Hymnal' LIMIT 1;
--   Song 3: 426 - Throw Out the Life-Line
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230521', 'song', 3 FROM songs WHERE normalized_name = 'throw out the life line' AND number = '426' AND book = 'Becky Sipla Music' LIMIT 1;
--   Song 4: 520 - O How I Love Jesus
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230521', 'song', 4 FROM songs WHERE normalized_name = 'o how i love jesus' AND number = '520' AND book = 'Hymnal' LIMIT 1;

-- May 28, 2023
--   Song 1: 072 - Thou Art Worthy
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230528', 'song', 1 FROM songs WHERE normalized_name = 'thou art worthy' AND number = '072' AND book = 'Hymnal' LIMIT 1;
--   Song 2: 486 - My Country Tis of Thee
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230528', 'song', 2 FROM songs WHERE normalized_name = 'my country tis of thee' AND number = '486' AND book = 'Hymnal' LIMIT 1;
--   Song 3: 521 - Ive Found a Friend
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230528', 'song', 3 FROM songs WHERE normalized_name = 'ive found a friend' AND number = '521' AND book = 'Hymnal' LIMIT 1;

-- Jun 4, 2023
--   Song 1: 205 - He Included Me
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230604', 'song', 1 FROM songs WHERE normalized_name = 'he included me' AND number = '205' AND book = 'Becky Sipla Music' LIMIT 1;
--   Song 2: 267 - Burdens Are Lifted At Calvary
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230604', 'song', 2 FROM songs WHERE normalized_name = 'burdens are lifted at calvary' AND number = '267' AND book = 'Hymnal' LIMIT 1;
--   Song 3: 522 - My Savior's Love
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230604', 'song', 3 FROM songs WHERE normalized_name = 'my saviors love' AND number = '522' AND book = 'Hymnal' LIMIT 1;
--   Song 4: 534 - Love Found A Way
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230604', 'song', 4 FROM songs WHERE normalized_name = 'love found a way' AND number = '534' AND book = 'Hymnal' LIMIT 1;

-- Jun 11, 2023
--   Song 1: 523 - Saved Saved
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230611', 'song', 1 FROM songs WHERE normalized_name = 'saved saved' AND number = '523' AND book = 'Hymnal' LIMIT 1;
--   Song 2: 535 - Satisfied
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230611', 'song', 2 FROM songs WHERE normalized_name = 'satisfied' AND number = '535' AND book = 'Hymnal' LIMIT 1;
--   Song 3: He is Able to Deliver Thee
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230611', 'song', 3 FROM songs WHERE normalized_name = 'he is able to deliver thee' AND book = 'Becky Sipla Music' LIMIT 1;

-- Jun 18, 2023
--   Song 1: 524 - At Calvary
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230618', 'song', 1 FROM songs WHERE normalized_name = 'at calvary' AND number = '524' AND book = 'Hymnal' LIMIT 1;
--   Song 2: Faith of our Fathers
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230618', 'song', 2 FROM songs WHERE normalized_name = 'faith of our fathers' AND book = 'Becky Sipla Music' LIMIT 1;
--   Song 3: Have I Done My Best for Jesus
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230618', 'song', 3 FROM songs WHERE normalized_name = 'have i done my best for jesus' AND book = 'Becky Sipla Music' LIMIT 1;

-- Jun 25, 2023
--   Song 1: 525 - Heaven Came Down
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230625', 'song', 1 FROM songs WHERE normalized_name = 'heaven came down' AND number = '525' AND book = 'Hymnal' LIMIT 1;
--   Song 2: 533 - I Believe in Miracles
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230625', 'song', 2 FROM songs WHERE normalized_name = 'i believe in miracles' AND number = '533' AND book = 'Hymnal' LIMIT 1;
--   Song 3: The Bible Stands
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230625', 'song', 3 FROM songs WHERE normalized_name = 'the bible stands' AND book = 'Becky Sipla Music' LIMIT 1;

-- Jul 2, 2023
--   Song 1: 224 - When I Survey The Wondrous Cross
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230702', 'song', 1 FROM songs WHERE normalized_name = 'when i survey the wondrous cross' AND number = '224' AND book = 'Hymnal' LIMIT 1;
--   Song 2: 526 - Victory in Jesus
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230702', 'song', 2 FROM songs WHERE normalized_name = 'victory in jesus' AND number = '526' AND book = 'Hymnal' LIMIT 1;
--   Song 3: The Star-Spangled Banner
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230702', 'song', 3 FROM songs WHERE normalized_name = 'the star spangled banner' AND book = 'Becky Sipla Music' LIMIT 1;

-- Jul 9, 2023
--   Song 1: 528 - Without Him
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230709', 'song', 1 FROM songs WHERE normalized_name = 'without him' AND number = '528' AND book = 'Hymnal' LIMIT 1;
--   Song 2: 537 - New Life
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230709', 'song', 2 FROM songs WHERE normalized_name = 'new life' AND number = '537' AND book = 'Hymnal' LIMIT 1;
--   Song 3: The Name of Jesus
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230709', 'song', 3 FROM songs WHERE normalized_name = 'the name of jesus' AND book = 'Becky Sipla Music' LIMIT 1;

-- Jul 16, 2023
--   Song 1: 529 - No Not One
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230716', 'song', 1 FROM songs WHERE normalized_name = 'no not one' AND number = '529' AND book = 'Hymnal' LIMIT 1;
--   Song 2: 538 - Now I Belong to Jesus
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230716', 'song', 2 FROM songs WHERE normalized_name = 'now i belong to jesus' AND number = '538' AND book = 'Hymnal' LIMIT 1;
--   Song 3: We Have an Anchor
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230716', 'song', 3 FROM songs WHERE normalized_name = 'we have an anchor' AND book = 'Becky Sipla Music' LIMIT 1;

-- Jul 23, 2023
--   Song 1: 531 - What A Wonderful Savior
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230723', 'song', 1 FROM songs WHERE normalized_name = 'what a wonderful savior' AND number = '531' AND book = 'Hymnal' LIMIT 1;
--   Song 2: Let the Lower Lights be Burning
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230723', 'song', 2 FROM songs WHERE normalized_name = 'let the lower lights be burning' AND book = 'Becky Sipla Music' LIMIT 1;
--   Song 3: The Lily of the Valley
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230723', 'song', 3 FROM songs WHERE normalized_name = 'the lily of the valley' AND book = 'Becky Sipla Music' LIMIT 1;

-- Jul 30, 2023
--   Song 1: 532 - He Touched Me
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230730', 'song', 1 FROM songs WHERE normalized_name = 'he touched me' AND number = '532' AND book = 'Hymnal' LIMIT 1;
--   Song 2: 539 - I Will Sing of My Redeemer
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230730', 'song', 2 FROM songs WHERE normalized_name = 'i will sing of my redeemer' AND number = '539' AND book = 'Hymnal' LIMIT 1;
--   Song 3: Lavish Love, Abundant Beauty
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230730', 'song', 3 FROM songs WHERE normalized_name = 'lavish love abundant beauty' AND book = 'Becky Sipla Music' LIMIT 1;

-- Aug 20, 2023
--   Song 1: 035 - Come Thou Fount
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230820', 'song', 1 FROM songs WHERE normalized_name = 'come thou fount' AND number = '035' AND book = 'Hymnal' LIMIT 1;
--   Song 2: 543 - Through It All
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230820', 'song', 2 FROM songs WHERE normalized_name = 'through it all' AND number = '543' AND book = 'Hymnal' LIMIT 1;
--   Song 3: What a Gathering
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230820', 'song', 3 FROM songs WHERE normalized_name = 'what a gathering' AND book = 'Becky Sipla Music' LIMIT 1;

-- Aug 27, 2023
--   Song 1: 037 - All Creatures of Our God and King
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230827', 'song', 1 FROM songs WHERE normalized_name = 'all creatures of our god and king' AND number = '037' AND book = 'Hymnal' LIMIT 1;
--   Song 2: 301 - Softly and Tenderly
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230827', 'song', 2 FROM songs WHERE normalized_name = 'softly and tenderly' AND number = '301' AND book = 'Hymnal' LIMIT 1;
--   Song 3: 545 - He Lifted Me
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230827', 'song', 3 FROM songs WHERE normalized_name = 'he lifted me' AND number = '545' AND book = 'Hymnal' LIMIT 1;

-- Sep 3, 2023
--   Song 1: 041 - Holy, Holy, Holy
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230903', 'song', 1 FROM songs WHERE normalized_name = 'holy holy holy' AND number = '041' AND book = 'Hymnal' LIMIT 1;
--   Song 2: 217 - Why
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230903', 'song', 2 FROM songs WHERE normalized_name = 'why' AND number = '217' AND book = 'Hymnal' LIMIT 1;
--   Song 3: Yesterday, Today, Forever
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20230903', 'song', 3 FROM songs WHERE normalized_name = 'yesterday today forever' AND book = 'Becky Sipla Music' LIMIT 1;

-- Oct 8, 2023
--   Song 1: 552 - In My Heart There Rings a Melody
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20231008', 'song', 1 FROM songs WHERE normalized_name = 'in my heart there rings a melody' AND number = '552' AND book = 'Hymnal' LIMIT 1;
--   Song 2: Its Just Like His Great Love
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20231008', 'song', 2 FROM songs WHERE normalized_name = 'its just like his great love' AND book = 'Becky Sipla Music' LIMIT 1;
--   Song 3: Prince Of Peace
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20231008', 'song', 3 FROM songs WHERE normalized_name = 'prince of peace' AND book = 'Choruses' LIMIT 1;

-- Oct 15, 2023
--   Song 1: 051 - Holy Holy
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20231015', 'song', 1 FROM songs WHERE normalized_name = 'holy holy' AND number = '051' AND book = 'Hymnal' LIMIT 1;
--   Song 2: 553 - Isnt The Love Of Jesus Something Wonderful
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20231015', 'song', 2 FROM songs WHERE normalized_name = 'isnt the love of jesus something wonderful' AND number = '553' AND book = 'Hymnal' LIMIT 1;
--   Song 3: Bringing in the Sheaves
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20231015', 'song', 3 FROM songs WHERE normalized_name = 'bringing in the sheaves' AND book = 'Becky Sipla Music' LIMIT 1;

-- Oct 22, 2023
--   Song 1: 053 - Joyful Joyful We Adore Thee
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20231022', 'song', 1 FROM songs WHERE normalized_name = 'joyful joyful we adore thee' AND number = '053' AND book = 'Hymnal' LIMIT 1;
--   Song 2: 558 - Springs of Living Water
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20231022', 'song', 2 FROM songs WHERE normalized_name = 'springs of living water' AND number = '558' AND book = 'Hymnal' LIMIT 1;
--   Song 3: The Lily of the Valley
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20231022', 'song', 3 FROM songs WHERE normalized_name = 'the lily of the valley' AND book = 'Becky Sipla Music' LIMIT 1;

-- Oct 29, 2023
--   Song 1: 054 - Great is Thy Faithfulness
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20231029', 'song', 1 FROM songs WHERE normalized_name = 'great is thy faithfulness' AND number = '054' AND book = 'Hymnal' LIMIT 1;
--   Song 2: 555 - Jesus Loves Me
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20231029', 'song', 2 FROM songs WHERE normalized_name = 'jesus loves me' AND number = '555' AND book = 'Hymnal' LIMIT 1;
--   Song 3: The Name of Jesus
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20231029', 'song', 3 FROM songs WHERE normalized_name = 'the name of jesus' AND book = 'Becky Sipla Music' LIMIT 1;

-- Nov 5, 2023
--   Song 1: 554 - The Old Rugged Cross
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20231105', 'song', 1 FROM songs WHERE normalized_name = 'the old rugged cross' AND number = '554' AND book = 'Hymnal' LIMIT 1;
--   Song 2: 557 - Redeemed
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20231105', 'song', 2 FROM songs WHERE normalized_name = 'redeemed' AND number = '557' AND book = 'Hymnal' LIMIT 1;
--   Song 3: Be Thou Exalted
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20231105', 'song', 3 FROM songs WHERE normalized_name = 'be thou exalted' AND book = 'Becky Sipla Music' LIMIT 1;

-- Dec 10, 2023
--   Song 1: 169 - Thou Didst Leave Thy Throne
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20231210', 'song', 1 FROM songs WHERE normalized_name = 'thou didst leave thy throne' AND number = '169' AND book = 'Hymnal' LIMIT 1;
--   Song 2: 182 - Angels from the Realms of Glory
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20231210', 'song', 2 FROM songs WHERE normalized_name = 'angels from the realms of glory' AND number = '182' AND book = 'Hymnal' LIMIT 1;
--   Song 3: 194 - As With Gladness Men Of Old
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20231210', 'song', 3 FROM songs WHERE normalized_name = 'as with gladness men of old' AND number = '194' AND book = 'Hymnal' LIMIT 1;

-- Dec 31, 2023
--   Song 1: 114 - Another Year is Dawning
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20231231', 'song', 1 FROM songs WHERE normalized_name = 'another year is dawning' AND number = '114' AND book = 'Hymnal' LIMIT 1;
--   Song 2: 172 - Come, Thou Long-Expected Jesus
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20231231', 'song', 2 FROM songs WHERE normalized_name = 'come thou long expected jesus' AND number = '172' AND book = 'Hymnal' LIMIT 1;
--   Song 3: 179 - Angels We Have Heard On High
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20231231', 'song', 3 FROM songs WHERE normalized_name = 'angels we have heard on high' AND number = '179' AND book = 'Hymnal' LIMIT 1;

-- Jan 7, 2024
--   Song 1: 227 - At the Cross
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20240107', 'song', 1 FROM songs WHERE normalized_name = 'at the cross' AND number = '227' AND book = 'Hymnal' LIMIT 1;
--   Song 2: 559 - Sunshine in My Soul
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20240107', 'song', 2 FROM songs WHERE normalized_name = 'sunshine in my soul' AND number = '559' AND book = 'Hymnal' LIMIT 1;
--   Song 3: Lift Up, Lift Up Your Voices Now
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20240107', 'song', 3 FROM songs WHERE normalized_name = 'lift up lift up your voices now' AND book = 'Becky Sipla Music' LIMIT 1;

-- Jan 14, 2024
--   Song 1: 057 - Our Great Savior
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20240114', 'song', 1 FROM songs WHERE normalized_name = 'our great savior' AND number = '057' AND book = 'Hymnal' LIMIT 1;
--   Song 2: 561 - Since Jesus Came into My Heart
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20240114', 'song', 2 FROM songs WHERE normalized_name = 'since jesus came into my heart' AND number = '561' AND book = 'Hymnal' LIMIT 1;
--   Song 3: Prince of Peace
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20240114', 'song', 3 FROM songs WHERE normalized_name = 'prince of peace' AND book = 'Becky Sipla Music' LIMIT 1;

-- Jan 21, 2024
--   Song 1: 058 - Fairest Lord Jesus
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20240121', 'song', 1 FROM songs WHERE normalized_name = 'fairest lord jesus' AND number = '058' AND book = 'Hymnal' LIMIT 1;
--   Song 2: 252 - Christ Returneth
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20240121', 'song', 2 FROM songs WHERE normalized_name = 'christ returneth' AND number = '252' AND book = 'Hymnal' LIMIT 1;
--   Song 3: 563 - All That Thrills My Soul
INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) SELECT id, '20240121', 'song', 3 FROM songs WHERE normalized_name = 'all that thrills my soul' AND number = '563' AND book = 'Hymnal' LIMIT 1;
