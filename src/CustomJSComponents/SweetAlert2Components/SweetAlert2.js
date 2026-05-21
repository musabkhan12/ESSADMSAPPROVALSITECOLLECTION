
import Swal from 'sweetalert2'
 export  const fireAlert = () => {
        Swal.fire({
            title: 'Do you want to save',
            showConfirmButton: true,
            showCancelButton: true,
            confirmButtonText: "Save",
            cancelButtonText: "Cancel",
            icon: 'warning'
        }
        ).then((result) => {
            /* Read more about isConfirmed, isDenied below */
            if (result.isConfirmed) {

                Swal.fire('Nice to meet you', '', 'success');

            } else
                Swal.fire(' Cancelled', '', 'error')

        })
    }
  