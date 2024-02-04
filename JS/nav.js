
var content = document.querySelector('.content');
var navbarToggle = document.querySelector('.navbar-toggle');
function toggleNavbar() {
    var navbar = document.querySelector('.navbar');
 

    if (navbar.style.left === '0px') {
        closeNavbar();
    } else {
        openNavbar();
    }
}

function openNavbar() {
    var navbar = document.querySelector('.navbar');
 
    var navbarToggle = document.querySelector('.navbar-toggle');

    navbar.style.left = '0';

    navbarToggle.style.left = '170px';
}

function closeNavbar() {
    var navbar = document.querySelector('.navbar');

    var navbarToggle = document.querySelector('.navbar-toggle');

    navbar.style.left = '-200px';
  
    navbarToggle.style.left = '0';
}