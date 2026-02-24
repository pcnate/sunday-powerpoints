/**
 * Generates a Kdenlive project XML file for sermon video editing.
 *
 * Produces an MLT XML document with:
 * - OBS recording (MKV, optional): 1 video track (desktop capture) + 3 audio tracks (desktop, mic, soundboard)
 * - Camera recording (MP4): 1 video track + 1 audio track
 * - 5 audio tracks + 5 video tracks (matching standard Kdenlive layout)
 * - Render URL pre-set to Vids/YYYYMMDD-production.mp4
 */


/**
 * Options for generating a Kdenlive project.
 */
interface KdenliveOptions {
  sundayDate: string;
  rootPath: string;
  obsFile?: string;
  cameraFile: string;
}


/**
 * Generate a unique UUID v4 string.
 */
function uuid(): string {
  return '{xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx}'.replace( /[xy]/g, ( c ) => {
    const r = ( Math.random() * 16 ) | 0;
    const v = c === 'x' ? r : ( r & 0x3 ) | 0x8;
    return v.toString( 16 );
  });
}


/**
 * Format a date string for the Kdenlive file.
 *
 * @param sundayDate - YYYYMMDD string
 * @returns formatted date like "2026-02-22"
 */
function formatDate( sundayDate: string ): string {
  return `${ sundayDate.slice( 0, 4 ) }-${ sundayDate.slice( 4, 6 ) }-${ sundayDate.slice( 6, 8 ) }`;
}


/**
 * Build XML for an audio tractor (paired playlists with hide="video" + standard filters).
 *
 * @param tractorId - tractor element ID
 * @param playlistAId - first playlist ID
 * @param playlistBId - second playlist ID
 * @param filterId - starting filter ID number
 * @param trackName - display name for the track
 * @param duration - total duration timecode
 * @returns XML string and next filter ID
 */
function audioTractor(
  tractorId: string,
  playlistAId: string,
  playlistBId: string,
  filterId: number,
  trackName: string,
  duration: string,
): { xml: string; nextFilterId: number } {
  const xml = ` <tractor id="${ tractorId }" in="00:00:00.000" out="${ duration }">
  <property name="kdenlive:audio_track">1</property>
  <property name="kdenlive:trackheight">64</property>
  <property name="kdenlive:timeline_active">1</property>
  <property name="kdenlive:collapsed">28</property>
  <property name="kdenlive:track_name">${ trackName }</property>
  <track hide="video" producer="${ playlistAId }"/>
  <track hide="video" producer="${ playlistBId }"/>
  <filter id="filter${ filterId }">
   <property name="window">75</property>
   <property name="max_gain">20dB</property>
   <property name="mlt_service">volume</property>
   <property name="internal_added">237</property>
   <property name="disable">1</property>
  </filter>
  <filter id="filter${ filterId + 1 }">
   <property name="channel">-1</property>
   <property name="mlt_service">panner</property>
   <property name="internal_added">237</property>
   <property name="start">0.5</property>
   <property name="disable">1</property>
  </filter>
  <filter id="filter${ filterId + 2 }">
   <property name="iec_scale">0</property>
   <property name="mlt_service">audiolevel</property>
   <property name="dbpeak">1</property>
   <property name="disable">1</property>
  </filter>
 </tractor>`;

  return { xml, nextFilterId: filterId + 3 };
}


/**
 * Build XML for a video tractor (paired playlists with hide="audio").
 *
 * @param tractorId - tractor element ID
 * @param playlistAId - first playlist ID
 * @param playlistBId - second playlist ID
 * @param trackName - display name for the track
 * @param duration - total duration timecode
 * @returns XML string
 */
function videoTractor(
  tractorId: string,
  playlistAId: string,
  playlistBId: string,
  trackName: string,
  duration: string,
): string {
  return ` <tractor id="${ tractorId }" in="00:00:00.000" out="${ duration }">
  <property name="kdenlive:trackheight">64</property>
  <property name="kdenlive:timeline_active">1</property>
  <property name="kdenlive:collapsed">28</property>
  <property name="kdenlive:track_name">${ trackName }</property>
  <track hide="audio" producer="${ playlistAId }"/>
  <track hide="audio" producer="${ playlistBId }"/>
 </tractor>`;
}


/**
 * Build a chain element for a media clip (audio or video reference).
 *
 * @param id - chain element ID
 * @param resource - relative file path
 * @param kdenliveId - clip ID in the project bin
 * @param controlUuid - shared UUID for clips from the same source file
 * @param audioIndex - audio stream index (1-based for the stream selector)
 * @param astream - audio stream index (0-based for astream property)
 * @param activeStreams - semicolon-separated active stream list (e.g., "1;2;3")
 * @param duration - duration timecode
 * @param testAudio - 0 to include audio, 1 to exclude
 * @param testImage - 0 to include video, 1 to exclude
 * @returns XML string
 */
function chain(
  id: string,
  resource: string,
  kdenliveId: number,
  controlUuid: string,
  audioIndex: number,
  astream: number,
  activeStreams: string,
  duration: string,
  testAudio: number,
  testImage: number,
): string {
  return ` <chain id="${ id }" out="${ duration }">
  <property name="length">2147483647</property>
  <property name="eof">pause</property>
  <property name="resource">${ resource }</property>
  <property name="mlt_service">avformat-novalidate</property>
  <property name="seekable">1</property>
  <property name="audio_index">${ audioIndex }</property>
  <property name="video_index">0</property>
  <property name="vstream">0</property>
  <property name="astream">${ astream }</property>
  <property name="kdenlive:folderid">-1</property>
  <property name="kdenlive:id">${ kdenliveId }</property>
  <property name="kdenlive:control_uuid">${ controlUuid }</property>
  <property name="mute_on_pause">0</property>
  <property name="kdenlive:active_streams">${ activeStreams }</property>
  <property name="kdenlive:clip_type">0</property>
  <property name="set.test_audio">${ testAudio }</property>
  <property name="set.test_image">${ testImage }</property>
 </chain>`;
}


/**
 * Generate a complete Kdenlive/MLT project XML string.
 *
 * Track layout (bottom to top in Kdenlive):
 * - A1: Camera audio
 * - A2: OBS Track 1 (desktop audio) — empty if no OBS
 * - A3: OBS Track 2 (mic) — empty if no OBS
 * - A4: OBS Track 3 (soundboard) — empty if no OBS
 * - A5: Empty spare audio track
 * - V1: Camera video
 * - V2: OBS desktop video — empty if no OBS
 * - V3-V5: Empty spare video tracks
 *
 * @param options - generation options
 * @returns complete MLT XML string
 */
export function generateKdenlive( options: KdenliveOptions ): string {
  const { sundayDate, rootPath, obsFile, cameraFile } = options;
  const hasObs = !!obsFile;
  const dateStr = formatDate( sundayDate );
  const sequenceUuid = uuid();
  const obsUuid = uuid();
  const cameraUuid = uuid();
  const documentId = Date.now().toString();

  // Default duration placeholder — user will edit in Kdenlive
  const dur = '01:30:00.000';

  const parts: string[] = [];

  // === XML header + profile ===
  parts.push( `<?xml version='1.0' encoding='utf-8'?>` );
  parts.push( `<mlt LC_NUMERIC="en_US.UTF-8" producer="main_bin" root="${ rootPath }" version="7.33.0">` );
  parts.push( ` <profile colorspace="709" description="HD 1080p 29.97 fps" display_aspect_den="9" display_aspect_num="16" frame_rate_den="1001" frame_rate_num="30000" height="1080" progressive="1" sample_aspect_den="1" sample_aspect_num="1" width="1920"/>` );

  // === producer0: black background ===
  parts.push( ` <producer id="producer0" in="00:00:00.000" out="${ dur }">
  <property name="length">2147483647</property>
  <property name="eof">continue</property>
  <property name="resource">black</property>
  <property name="aspect_ratio">1</property>
  <property name="mlt_service">color</property>
  <property name="kdenlive:playlistid">black_track</property>
  <property name="mlt_image_format">rgba</property>
  <property name="set.test_audio">0</property>
 </producer>` );

  // === Clip chains ===
  // OBS clip IDs: kdenlive:id=5 (matching reference pattern)
  // Camera clip IDs: kdenlive:id=6

  // Chain 0: Camera audio (audio_index=1, astream=0)
  parts.push( chain( 'chain0', cameraFile, 6, cameraUuid, 1, 0, '', dur, 0, 1 ) );

  if ( hasObs ) {
    // Chain 1: OBS audio Track 1 — desktop (audio_index=1, astream=0)
    parts.push( chain( 'chain1', obsFile!, 5, obsUuid, 1, 0, '1;2;3', dur, 0, 1 ) );

    // Chain 2: OBS audio Track 2 — mic (audio_index=2, astream=1)
    parts.push( chain( 'chain2', obsFile!, 5, obsUuid, 2, 1, '1;2;3', dur, 0, 1 ) );

    // Chain 3: OBS audio Track 3 — soundboard (audio_index=3, astream=2)
    parts.push( chain( 'chain3', obsFile!, 5, obsUuid, 3, 2, '1;2;3', dur, 0, 1 ) );
  }

  // Chain 4: Camera video (video only)
  parts.push( chain( 'chain4', cameraFile, 6, cameraUuid, 1, 0, '', dur, 1, 0 ) );

  if ( hasObs ) {
    // Chain 5: OBS video (video only)
    parts.push( chain( 'chain5', obsFile!, 5, obsUuid, 1, 0, '1;2;3', dur, 1, 0 ) );
  }

  // === Audio track playlists + tractors ===
  // A1: Camera audio
  parts.push( ` <playlist id="playlist0">
  <entry in="00:00:00.000" out="${ dur }" producer="chain0">
   <property name="kdenlive:id">6</property>
  </entry>
 </playlist>` );
  parts.push( ` <playlist id="playlist1"/>` );

  let fId = 0;
  let result = audioTractor( 'tractor0', 'playlist0', 'playlist1', fId, 'Camera Audio', dur );
  parts.push( result.xml );
  fId = result.nextFilterId;

  // A2: OBS Track 1 (desktop audio) — empty if no OBS
  if ( hasObs ) {
    parts.push( ` <playlist id="playlist2">
  <entry in="00:00:00.000" out="${ dur }" producer="chain1">
   <property name="kdenlive:id">5</property>
  </entry>
 </playlist>` );
  } else {
    parts.push( ` <playlist id="playlist2"/>` );
  }
  parts.push( ` <playlist id="playlist3"/>` );

  result = audioTractor( 'tractor1', 'playlist2', 'playlist3', fId, 'Desktop Audio', dur );
  parts.push( result.xml );
  fId = result.nextFilterId;

  // A3: OBS Track 2 (mic) — empty if no OBS
  if ( hasObs ) {
    parts.push( ` <playlist id="playlist4">
  <entry in="00:00:00.000" out="${ dur }" producer="chain2">
   <property name="kdenlive:id">5</property>
  </entry>
 </playlist>` );
  } else {
    parts.push( ` <playlist id="playlist4"/>` );
  }
  parts.push( ` <playlist id="playlist5"/>` );

  result = audioTractor( 'tractor2', 'playlist4', 'playlist5', fId, 'Mic', dur );
  parts.push( result.xml );
  fId = result.nextFilterId;

  // A4: OBS Track 3 (soundboard) — empty if no OBS
  if ( hasObs ) {
    parts.push( ` <playlist id="playlist6">
  <entry in="00:00:00.000" out="${ dur }" producer="chain3">
   <property name="kdenlive:id">5</property>
  </entry>
 </playlist>` );
  } else {
    parts.push( ` <playlist id="playlist6"/>` );
  }
  parts.push( ` <playlist id="playlist7"/>` );

  result = audioTractor( 'tractor3', 'playlist6', 'playlist7', fId, 'Soundboard', dur );
  parts.push( result.xml );
  fId = result.nextFilterId;

  // A5: Empty spare audio track
  parts.push( ` <playlist id="playlist8"/>` );
  parts.push( ` <playlist id="playlist9"/>` );

  result = audioTractor( 'tractor4', 'playlist8', 'playlist9', fId, 'Audio 5', dur );
  parts.push( result.xml );
  fId = result.nextFilterId;

  // === Video track playlists + tractors ===
  // V1: Camera video
  parts.push( ` <playlist id="playlist10">
  <entry in="00:00:00.000" out="${ dur }" producer="chain4">
   <property name="kdenlive:id">6</property>
  </entry>
 </playlist>` );
  parts.push( ` <playlist id="playlist11"/>` );
  parts.push( videoTractor( 'tractor5', 'playlist10', 'playlist11', 'Camera', dur ) );

  // V2: OBS desktop video — empty if no OBS
  if ( hasObs ) {
    parts.push( ` <playlist id="playlist12">
  <entry in="00:00:00.000" out="${ dur }" producer="chain5">
   <property name="kdenlive:id">5</property>
  </entry>
 </playlist>` );
  } else {
    parts.push( ` <playlist id="playlist12"/>` );
  }
  parts.push( ` <playlist id="playlist13"/>` );
  parts.push( videoTractor( 'tractor6', 'playlist12', 'playlist13', 'OBS Desktop', dur ) );

  // V3-V5: Empty spare video tracks
  for ( let i = 0; i < 3; i++ ) {
    const pA = `playlist${ 14 + i * 2 }`;
    const pB = `playlist${ 15 + i * 2 }`;
    const tId = `tractor${ 7 + i }`;
    parts.push( ` <playlist id="${ pA }"/>` );
    parts.push( ` <playlist id="${ pB }"/>` );
    parts.push( videoTractor( tId, pA, pB, `Video ${ i + 3 }`, dur ) );
  }

  // === Sequence tractor (the main timeline) ===
  // Track layout: producer0 (black) + 5 audio tractors + 5 video tractors = 11 tracks
  //
  // Group track indices (0-based from first tractor, skipping producer0):
  //   0: Camera Audio     ← Camera group
  //   1: Desktop Audio    ← OBS group
  //   2: Mic              ← OBS group
  //   3: Soundboard       ← OBS group
  //   4: Audio 5 (spare)
  //   5: Camera Video     ← Camera group
  //   6: OBS Desktop      ← OBS group
  //   7-9: spare video
  const groups: object[] = [];

  // OBS group: links OBS audio tracks (1-3) + OBS video track (6)
  if ( hasObs ) {
    groups.push({
      children: [
        { data: '1:0:-1', leaf: 'clip', type: 'Leaf' },
        { data: '2:0:-1', leaf: 'clip', type: 'Leaf' },
        { data: '3:0:-1', leaf: 'clip', type: 'Leaf' },
        { data: '6:0:-1', leaf: 'clip', type: 'Leaf' },
      ],
      type: 'Normal',
    });
  }

  // Camera group: links camera audio (0) + camera video (5)
  groups.push({
    children: [
      { data: '0:0:-1', leaf: 'clip', type: 'Leaf' },
      { data: '5:0:-1', leaf: 'clip', type: 'Leaf' },
    ],
    type: 'AVSplit',
  });

  const seqGroups = JSON.stringify( groups, null, 4 );
  const seqGuides = JSON.stringify( [] );

  parts.push( ` <tractor id="${ sequenceUuid }" in="00:00:00.000" out="${ dur }">
  <property name="kdenlive:uuid">${ sequenceUuid }</property>
  <property name="kdenlive:clipname">Sequence 1</property>
  <property name="kdenlive:sequenceproperties.hasAudio">1</property>
  <property name="kdenlive:sequenceproperties.hasVideo">1</property>
  <property name="kdenlive:sequenceproperties.activeTrack">6</property>
  <property name="kdenlive:sequenceproperties.tracksCount">10</property>
  <property name="kdenlive:sequenceproperties.documentuuid">${ sequenceUuid }</property>
  <property name="kdenlive:control_uuid">${ sequenceUuid }</property>
  <property name="kdenlive:duration">${ dur }</property>
  <property name="kdenlive:maxduration">2147483647</property>
  <property name="kdenlive:producer_type">17</property>
  <property name="kdenlive:id">3</property>
  <property name="kdenlive:clip_type">0</property>
  <property name="kdenlive:file_size">0</property>
  <property name="kdenlive:folderid">2</property>
  <property name="kdenlive:sequenceproperties.audioTarget">1</property>
  <property name="kdenlive:sequenceproperties.disablepreview">0</property>
  <property name="kdenlive:sequenceproperties.position">0</property>
  <property name="kdenlive:sequenceproperties.scrollPos">0</property>
  <property name="kdenlive:sequenceproperties.tracks">4</property>
  <property name="kdenlive:sequenceproperties.verticalzoom">1</property>
  <property name="kdenlive:sequenceproperties.videoTarget">6</property>
  <property name="kdenlive:sequenceproperties.zonein">0</property>
  <property name="kdenlive:sequenceproperties.zoneout">75</property>
  <property name="kdenlive:sequenceproperties.zoom">8</property>
  <property name="kdenlive:sequenceproperties.groups">${ seqGroups }
</property>
  <property name="kdenlive:sequenceproperties.guides">${ seqGuides }
</property>
  <track producer="producer0"/>
  <track producer="tractor0"/>
  <track producer="tractor1"/>
  <track producer="tractor2"/>
  <track producer="tractor3"/>
  <track producer="tractor4"/>
  <track producer="tractor5"/>
  <track producer="tractor6"/>
  <track producer="tractor7"/>
  <track producer="tractor8"/>
  <track producer="tractor9"/>` );

  // Audio mix transitions (tracks 1-5 → audio tractors)
  for ( let i = 0; i < 5; i++ ) {
    parts.push( `  <transition id="transition${ i }">
   <property name="a_track">0</property>
   <property name="b_track">${ i + 1 }</property>
   <property name="mlt_service">mix</property>
   <property name="kdenlive_id">mix</property>
   <property name="internal_added">237</property>
   <property name="always_active">1</property>
   <property name="accepts_blanks">1</property>
   <property name="sum">1</property>
  </transition>` );
  }

  // Video composite transitions (tracks 6-10 → video tractors)
  for ( let i = 0; i < 5; i++ ) {
    parts.push( `  <transition id="transition${ 5 + i }">
   <property name="a_track">0</property>
   <property name="b_track">${ 6 + i }</property>
   <property name="compositing">0</property>
   <property name="distort">0</property>
   <property name="rotate_center">0</property>
   <property name="mlt_service">qtblend</property>
   <property name="kdenlive_id">qtblend</property>
   <property name="internal_added">237</property>
   <property name="always_active">1</property>
  </transition>` );
  }

  // Master volume + panner filters on the sequence
  parts.push( `  <filter id="filter${ fId }">
   <property name="window">75</property>
   <property name="max_gain">20dB</property>
   <property name="mlt_service">volume</property>
   <property name="internal_added">237</property>
   <property name="disable">1</property>
  </filter>
  <filter id="filter${ fId + 1 }">
   <property name="channel">-1</property>
   <property name="mlt_service">panner</property>
   <property name="internal_added">237</property>
   <property name="start">0.5</property>
   <property name="disable">1</property>
  </filter>
 </tractor>` );
  fId += 2;

  // === Bin-only chains for the project bin (full clips without test flags) ===
  // Camera bin clip
  parts.push( ` <chain id="chain_cam_bin" out="${ dur }">
  <property name="length">2147483647</property>
  <property name="eof">pause</property>
  <property name="resource">${ cameraFile }</property>
  <property name="mlt_service">avformat-novalidate</property>
  <property name="seekable">1</property>
  <property name="audio_index">1</property>
  <property name="video_index">0</property>
  <property name="kdenlive:id">6</property>
  <property name="kdenlive:control_uuid">${ cameraUuid }</property>
  <property name="kdenlive:clip_type">0</property>
  <property name="kdenlive:folderid">-1</property>
 </chain>` );

  // OBS bin clip (only when OBS file present)
  if ( hasObs ) {
    parts.push( ` <chain id="chain_obs_bin" out="${ dur }">
  <property name="length">2147483647</property>
  <property name="eof">pause</property>
  <property name="resource">${ obsFile }</property>
  <property name="mlt_service">avformat-novalidate</property>
  <property name="seekable">1</property>
  <property name="audio_index">1</property>
  <property name="video_index">0</property>
  <property name="kdenlive:id">5</property>
  <property name="kdenlive:control_uuid">${ obsUuid }</property>
  <property name="kdenlive:active_streams">1;2;3</property>
  <property name="kdenlive:clip_type">0</property>
  <property name="kdenlive:folderid">-1</property>
 </chain>` );
  }

  // === main_bin ===
  const guidesCategories = JSON.stringify( [
    { color: '#9b59b6', comment: 'Category 1', index: 0 },
    { color: '#3daee9', comment: 'Category 2', index: 1 },
    { color: '#1abc9c', comment: 'Category 3', index: 2 },
    { color: '#1cdc9a', comment: 'Category 4', index: 3 },
    { color: '#c9ce3b', comment: 'Category 5', index: 4 },
    { color: '#fdbc4b', comment: 'Category 6', index: 5 },
    { color: '#f39c1f', comment: 'Category 7', index: 6 },
    { color: '#f47750', comment: 'Category 8', index: 7 },
    { color: '#da4453', comment: 'Category 9', index: 8 },
  ], null, 4 );

  parts.push( ` <playlist id="main_bin">
  <property name="kdenlive:folder.-1.2">Sequences</property>
  <property name="kdenlive:sequenceFolder">2</property>
  <property name="kdenlive:docproperties.audioChannels">2</property>
  <property name="kdenlive:docproperties.documentid">${ documentId }</property>
  <property name="kdenlive:docproperties.enableTimelineZone">0</property>
  <property name="kdenlive:docproperties.enableproxy">0</property>
  <property name="kdenlive:docproperties.generateproxy">0</property>
  <property name="kdenlive:docproperties.guidesCategories">${ guidesCategories }
</property>
  <property name="kdenlive:docproperties.kdenliveversion">25.04.3</property>
  <property name="kdenlive:docproperties.profile">atsc_1080p_2997</property>
  <property name="kdenlive:docproperties.seekOffset">30000</property>
  <property name="kdenlive:docproperties.sessionid">${ uuid() }</property>
  <property name="kdenlive:docproperties.uuid">${ sequenceUuid }</property>
  <property name="kdenlive:docproperties.version">1.1</property>
  <property name="kdenlive:expandedFolders">2</property>
  <property name="kdenlive:binZoom">4</property>
  <property name="kdenlive:extraBins">project_bin:-1:0</property>
  <property name="kdenlive:documentnotes"/>
  <property name="kdenlive:documentnotesversion">2</property>
  <property name="kdenlive:docproperties.opensequences">${ sequenceUuid }</property>
  <property name="kdenlive:docproperties.activetimeline">${ sequenceUuid }</property>
  <property name="kdenlive:docproperties.rendercategory">Generic (HD for web, mobile devices...)</property>
  <property name="kdenlive:docproperties.rendercustomquality">-1</property>
  <property name="kdenlive:docproperties.renderendguide">-1</property>
  <property name="kdenlive:docproperties.renderexportaudio">0</property>
  <property name="kdenlive:docproperties.renderfullcolorrange">0</property>
  <property name="kdenlive:docproperties.rendermode">0</property>
  <property name="kdenlive:docproperties.renderplay">0</property>
  <property name="kdenlive:docproperties.renderpreview">0</property>
  <property name="kdenlive:docproperties.renderprofile">MP4-H264/AAC</property>
  <property name="kdenlive:docproperties.renderrescale">0</property>
  <property name="kdenlive:docproperties.renderspeed">6</property>
  <property name="kdenlive:docproperties.renderstartguide">-1</property>
  <property name="kdenlive:docproperties.rendertcoverlay">0</property>
  <property name="kdenlive:docproperties.rendertctype">-1</property>
  <property name="kdenlive:docproperties.rendertwopass">0</property>
  <property name="kdenlive:docproperties.renderurl">${ sundayDate }-production.mp4</property>
  <property name="xml_retain">1</property>
  <entry in="00:00:00.000" out="00:00:00.000" producer="${ sequenceUuid }"/>
  <entry in="00:00:00.000" out="${ dur }" producer="chain_cam_bin"/>${ hasObs ? `
  <entry in="00:00:00.000" out="${ dur }" producer="chain_obs_bin"/>` : '' }
 </playlist>` );

  // === Top-level project tractor ===
  parts.push( ` <tractor id="tractor_project" in="00:00:00.000" out="${ dur }">
  <property name="kdenlive:projectTractor">1</property>
  <track in="00:00:00.000" out="${ dur }" producer="${ sequenceUuid }"/>
 </tractor>` );

  parts.push( `</mlt>` );

  return parts.join( '\n' );
}
