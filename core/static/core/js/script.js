// Main script for RENEE Warehouse
console.log("Welcome to RENEE Warehouse.");

// Helper to create toasts programmatically
window.createToast = function(message, type) {
    const container = document.getElementById('toast-container');
    if (!container) return;

    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;

    const msgSpan = document.createElement('span');
    msgSpan.innerText = message;

    const closeBtn = document.createElement('button');
    closeBtn.className = 'toast-close';
    closeBtn.innerHTML = '&times;';
    closeBtn.onclick = function() { toast.remove(); };

    toast.appendChild(msgSpan);
    toast.appendChild(closeBtn);
    container.appendChild(toast);

    setTimeout(() => {
        if(toast && toast.parentElement) {
            toast.style.transition = "opacity 0.3s ease-out";
            toast.style.opacity = "0";
            setTimeout(() => toast.remove(), 300);
        }
    }, 5000);
};

// Auto-dismiss server-rendered toasts after 5 seconds
document.addEventListener("DOMContentLoaded", () => {
    const toasts = document.querySelectorAll(".toast");
    toasts.forEach((toast) => {
        setTimeout(() => {
            if(toast && toast.parentElement) {
                toast.style.transition = "opacity 0.3s ease-out";
                toast.style.opacity = "0";
                setTimeout(() => toast.remove(), 300);
            }
        }, 5000);
    });
});
