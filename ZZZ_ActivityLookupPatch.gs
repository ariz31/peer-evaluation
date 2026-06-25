function getOrCreateActivity_(ak,ci,gk,name){
  var rows=activities_();
  var exact=rows.find(function(x){return x.activityKey===ak&&x.classId===ci&&x.groupKey===gk});
  if(exact)return exact;
  var shared=rows.find(function(x){return x.activityKey===ak&&x.classId===ci&&x.groupKey===CONFIG.ALL});
  if(shared)return shared;
  return upsertActivity_({activityKey:ak,classId:ci,groupKey:gk,activityName:name||ak,status:'Active',allowRegistration:true,regOpen:'',regClose:'',evalOpen:'',evalClose:''});
}
