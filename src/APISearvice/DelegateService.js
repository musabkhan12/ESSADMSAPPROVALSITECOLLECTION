import Swal from 'sweetalert2';

export const getDelegateList = async (_sp, isSuperAdmin) => {
    let arr = [];
    try {
        const currentUser = await _sp.web.currentUser();
        
        let query = _sp.web.lists.getByTitle("DelegateList").items
            .select("*,DelegateName/ID,DelegateName/Title,DelegateName/EMail,ActingFor/ID,ActingFor/Title,ActingFor/EMail")
            .expand("ActingFor,DelegateName")
            .orderBy("Created", false);

        if (isSuperAdmin === "Yes") {
            // Get all items for super admin - FIX: Remove .getAll()
            arr = await query();
        } else {
            // Filter for non-admin users
            arr = await query
                .filter(`AuthorId eq '${currentUser.Id}'`)();
        }
        
        console.log("Delegate list:", arr);
        return arr;
    } catch (error) {
        console.log("Error fetching data: ", error);
        return [];
    }
};

export const DeleteDelegateAPI = async (_sp, id) => {
    try {
        await _sp.web.lists.getByTitle('DelegateList')
            .items.getById(id)
            .delete();
        return true;
    } catch (error) {
        console.log('Error deleting item:', error);
        return false;
    }
};

export const addItem = async (itemData, _sp) => {
    try {
        const newItem = await _sp.web.lists.getByTitle('DelegateList')
            .items.add(itemData);
        return newItem;
    } catch (error) {
        console.log('Error adding item:', error);
        Swal.fire('Error', 'Failed to add item', 'error');
        return null;
    }
};

export const getDelegateByID = async (_sp, id) => {
    try {
        const res = await _sp.web.lists.getByTitle("DelegateList")
            .items.getById(id)
            .select("*,DelegateName/ID,DelegateName/Title,DelegateName/EMail,ActingFor/ID,ActingFor/Title,ActingFor/EMail")
            .expand("ActingFor,DelegateName")();
        
        const parsedValues = {
            ID: res.ID,
            StartDate: res.StartDate || "",
            EndDate: res.EndDate || "",
            Status: res.Status || "",
            DelegateName: res.DelegateName || "",
            ActingFor: res.ActingFor || "",
        };
        
        return [parsedValues];
    } catch (error) {
        console.log("Error fetching data: ", error);
        return [];
    }
};

export const updateItem = async (itemData, _sp, id) => {
    try {
        const updatedItem = await _sp.web.lists.getByTitle('DelegateList')
            .items.getById(id)
            .update(itemData);
        return updatedItem;
    } catch (error) {
        console.log('Error updating item:', error);
        return null;
    }
};