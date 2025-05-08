document.addEventListener('DOMContentLoaded', function() {
    // Add any custom JavaScript here
    console.log('PDF Converter Pro is ready!');

    // Example: Add confirmation for delete actions
    const deleteButtons = document.querySelectorAll('.btn-delete');
    deleteButtons.forEach(button => {
        button.addEventListener('click', function(e) {
            if (!confirm('Are you sure you want to delete this file?')) {
                e.preventDefault();
            }
        });
    });
});