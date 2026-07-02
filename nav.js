document.addEventListener('click',function(e){
  var t=e.target.closest('.nav-toggle');
  if(t){document.querySelector('.main-nav').classList.toggle('open');}
});
