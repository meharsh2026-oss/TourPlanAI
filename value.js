// =======================================
// TourPlan AI - VALUE Engine v1.0
// =======================================

const ValueEngine = {

    TARGET_TOTAL: 250000,

    update(rows){

        let total = 0;
        let lastWorkingRow = -1;

        // VALUE is column H (index 7)
        for(let i = 5; i < rows.length; i++){

            const row = rows[i];

            if(!row) continue;

            // Skip Sundays
            if(String(row[1]).toUpperCase() === "SUNDAY"){
                continue;
            }

            let value = Number(row[7]) || 0;

            total += value;

            lastWorkingRow = i;
        }

        if(lastWorkingRow === -1){
            return rows;
        }

        const difference = this.TARGET_TOTAL - total;

        rows[lastWorkingRow][7] =
            (Number(rows[lastWorkingRow][7]) || 0)
            + difference;

        return rows;

    }

};
