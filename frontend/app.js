function toggleSidebar() {
    document.getElementById('sidebar').classList.toggle('open');
    document.getElementById('sidebarOverlay').classList.toggle('active');
}

function showPage(pageId, buttonElement) {
    // Sabhi pages ko hide karein
    document.querySelectorAll('.page-section').forEach(sec => {
        sec.classList.remove('active');
    });

    // Selected page ko show karein
    document.getElementById('page-' + pageId).classList.add('active');

    // Sidebar active button update karein
    document.querySelectorAll('.nav-item').forEach(btn => {
        btn.classList.remove('active');
    });
    if (buttonElement) {
        buttonElement.classList.add('active');
        document.getElementById('pageTitle').innerText = buttonElement.innerText.trim();
    }

    // Mobile par sidebar automatically band karein
    if (window.innerWidth <= 768) {
        document.getElementById('sidebar').classList.remove('open');
        document.getElementById('sidebarOverlay').classList.remove('active');
    }
}

function navigateTo(pageId) {
    const buttons = document.querySelectorAll('.nav-item');
    buttons.forEach(btn => {
        if (btn.getAttribute('onclick').includes(`'${pageId}'`)) {
            showPage(pageId, btn);
        }
    });
}