function optionalCode_(v){v=text_(v,100).toUpperCase();return v||''}
function saveActivitySettingsOptional(d,p){d=d||{};d.groupKey=optionalCode_(d.groupKey)||CONFIG.ALL;return saveActivitySettings(d,p)}
