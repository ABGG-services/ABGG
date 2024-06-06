
function updateLogoScales() {
	
	var aelogos = document.querySelectorAll('.agenda-event.logo-image');
    // Find the smallest child height
	let aelogosminHeight = Array.prototype.reduce.call(aelogos, function(smallest, current) {
	  return current.offsetHeight < smallest ? current.offsetHeight : smallest;
	}, aelogos[0].offsetHeight); 
	
	aelogos.forEach(logoimg => {
		logoimg.style.setProperty('--height',aelogosminHeight+"px");
    });
	
	var celogos = document.querySelectorAll('.calendar-event.logo-image');
	let celogosminHeight = Array.prototype.reduce.call(celogos, function(smallest, current) {
	  return current.offsetHeight < smallest ? current.offsetHeight : smallest;
	}, celogos[0].offsetHeight); 

	celogos.forEach(logoimg => {
		logoimg.style.setProperty('--height',celogosminHeight+"px");
    });

}

// Call the function initially and whenever the window is resized
window.addEventListener('resize', updateLogoScales);
window.addEventListener('load', updateLogoScales); // In case images load after initial load