document.addEventListener('DOMContentLoaded', () => {
    const card = document.getElementById('tilt-card');
    const container = document.querySelector('.container');

    container.addEventListener('mousemove', (e) => {
        // Calculate position relative to the window since container is full screen
        const x = e.clientX;
        const y = e.clientY;
        
        const centerX = window.innerWidth / 2;
        const centerY = window.innerHeight / 2;
        
        // Calculate rotation based on cursor distance from center
        // Max rotation is 10 degrees at the edges
        const rotateX = ((y - centerY) / centerY) * -10; 
        const rotateY = ((x - centerX) / centerX) * 10;

        card.style.transform = `rotateX(${rotateX}deg) rotateY(${rotateY}deg)`;
    });

    container.addEventListener('mouseleave', () => {
        card.style.transform = 'rotateX(0) rotateY(0)';
    });
});
