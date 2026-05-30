document.addEventListener('DOMContentLoaded', () => {
    // 1. Scroll Reveal Animation using Intersection Observer
    const revealElements = document.querySelectorAll('.reveal');

    const revealOptions = {
        threshold: 0.15, // Trigger when 15% of the element is visible
        rootMargin: "0px 0px -50px 0px" // Slight offset so it triggers a bit earlier before bottom
    };

    const revealOnScroll = new IntersectionObserver(function(entries, observer) {
        entries.forEach(entry => {
            if (!entry.isIntersecting) {
                return;
            } else {
                entry.target.classList.add('active');
                // Optional: stop observing once revealed
                // observer.unobserve(entry.target); 
            }
        });
    }, revealOptions);

    revealElements.forEach(el => {
        revealOnScroll.observe(el);
    });

    // 2. Header Style on Scroll
    const header = document.getElementById('header');
    
    window.addEventListener('scroll', () => {
        if (window.scrollY > 50) {
            header.style.background = 'rgba(250, 249, 246, 0.96)';
            header.style.boxShadow = '0 4px 20px rgba(22, 101, 52, 0.06)';
            header.style.padding = '15px 50px';
        } else {
            header.style.background = 'rgba(250, 249, 246, 0.75)';
            header.style.boxShadow = 'none';
            header.style.padding = '20px 50px';
        }
    });

    // 3. Smooth scrolling for nav links
    document.querySelectorAll('a[href^="#"]').forEach(anchor => {
        anchor.addEventListener('click', function (e) {
            e.preventDefault();
            const targetId = this.getAttribute('href');
            if (targetId === '#') return;
            
            const targetElement = document.querySelector(targetId);
            if (targetElement) {
                targetElement.scrollIntoView({
                    behavior: 'smooth',
                    block: 'start'
                });
            }
        });
    });

    // 4. Tab Switching Logic for About Section
    const tabButtons = document.querySelectorAll('.tab-btn');
    const tabPanes = document.querySelectorAll('.tab-pane');
    
    tabButtons.forEach(btn => {
        btn.addEventListener('click', () => {
            const tabId = btn.getAttribute('data-tab');
            
            // Remove active from all buttons and panes
            tabButtons.forEach(b => b.classList.remove('active'));
            tabPanes.forEach(p => p.classList.remove('active'));
            
            // Activate clicked button and corresponding pane
            btn.classList.add('active');
            const targetPane = document.getElementById(tabId);
            if (targetPane) {
                targetPane.classList.add('active');
            }
        });
    });

    // 5. Image Zoom Modal Logic
    const modal = document.getElementById('image-modal');
    const modalImg = document.getElementById('modal-img');
    const modalTitle = document.getElementById('modal-caption-title');
    const modalDesc = document.getElementById('modal-caption-desc');
    const closeBtn = document.getElementById('modal-close-btn');
    const prevBtn = document.getElementById('modal-prev-btn');
    const nextBtn = document.getElementById('modal-next-btn');

    // Collect all zoomable cards in order of appearance
    const zoomableCards = [];
    
    // First, honor cards
    document.querySelectorAll('.honor-card').forEach(card => {
        const titleEl = card.querySelector('h3');
        const descEl = card.querySelector('p');
        zoomableCards.push({
            imgSrc: card.querySelector('img').src,
            title: titleEl ? titleEl.textContent : '',
            desc: descEl ? descEl.textContent : '',
            element: card
        });
    });

    // Then, celebrity photo cards
    document.querySelectorAll('.photo-card').forEach(card => {
        const titleEl = card.querySelector('h4');
        const descEl = card.querySelector('p');
        zoomableCards.push({
            imgSrc: card.querySelector('img').src,
            title: titleEl ? titleEl.textContent : '',
            desc: descEl ? descEl.textContent : '',
            element: card
        });
    });

    let currentPhotoIndex = 0;

    function openModal(index) {
        currentPhotoIndex = index;
        const data = zoomableCards[index];
        modalImg.src = data.imgSrc;
        modalTitle.textContent = data.title;
        modalDesc.textContent = data.desc;
        modal.classList.add('active');
        document.body.style.overflow = 'hidden'; // prevent scrolling behind modal
    }

    function closeModal() {
        modal.classList.remove('active');
        document.body.style.overflow = '';
    }

    function showNext() {
        let nextIndex = (currentPhotoIndex + 1) % zoomableCards.length;
        openModal(nextIndex);
    }

    function showPrev() {
        let prevIndex = (currentPhotoIndex - 1 + zoomableCards.length) % zoomableCards.length;
        openModal(prevIndex);
    }

    // Attach click events to the cards
    zoomableCards.forEach((cardData, index) => {
        cardData.element.addEventListener('click', (e) => {
            e.preventDefault();
            e.stopPropagation();
            openModal(index);
        });
    });

    // Modal Control Events
    closeBtn.addEventListener('click', closeModal);
    prevBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        showPrev();
    });
    nextBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        showNext();
    });

    // Click backdrop to close
    modal.addEventListener('click', (e) => {
        if (e.target === modal || e.target.classList.contains('modal-content-container')) {
            closeModal();
        }
    });

    // Keyboard navigation
    document.addEventListener('keydown', (e) => {
        if (!modal.classList.contains('active')) return;
        
        if (e.key === 'Escape') {
            closeModal();
        } else if (e.key === 'ArrowRight') {
            showNext();
        } else if (e.key === 'ArrowLeft') {
            showPrev();
        }
    });
});
