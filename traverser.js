const getColumnCount = (row) => {
  const columns = row.getElementsByTagName('td');
  let count = 0;

  for (const column of columns) {
    count += parseInt(column.getAttribute('colspan'));
  }
  return count;
}

(() => {
  const tbl = document.getElementById('tbl');
  const rows = tbl.getElementsByTagName('tr');
  const rowCount = rows.length;
  const columnCount = getColumnCount(rows[0]);

  let propertiesMatrix = new Array(rowCount)
    .fill({})
    .map(() => new Array(columnCount).fill({}));

  for (let ri = 0; ri < rowCount; ri++) {
    const columns = rows[ri].getElementsByTagName('td');
    let itr = 0;
    for (let ci = 0; ci < columns.length; ci++) {
      const rspan = parseInt(columns[ci].getAttribute('rowspan'));
      const cspan = parseInt(columns[ci].getAttribute('colspan'));

      if (ri === 2) console.log(`${ri} - ${rspan}, ${ci} - ${cspan}`);
      if (ri === 2) console.log(propertiesMatrix[ri - 1][itr], itr);
      while (ri > 0 &&
        propertiesMatrix[ri - 1][itr] &&
        propertiesMatrix[ri - 1][itr]['w:vMerge'] &&
        ri < propertiesMatrix[ri - 1][itr]['endAt']) {
        propertiesMatrix[ri][itr] = { ...propertiesMatrix[ri - 1][itr], 'w:vMerge': true };
        itr += (propertiesMatrix[ri - 1][itr]['w:gridSpan'] ? propertiesMatrix[ri - 1][itr]['w:gridSpan'] : 1);
      }

      if (rspan && cspan) {
        propertiesMatrix[ri][itr] = { ...propertiesMatrix[ri][itr], 'w:vMerge': "restart", 'endAt': ri + rspan, 'w:gridSpan': cspan };
        itr += cspan;
      } else if (rspan) {
        propertiesMatrix[ri][itr] = { ...propertiesMatrix[ri][itr], 'w:vMerge': "restart", 'endAt': ri + rspan };
        if (!cspan) itr++;
      } else if (cspan) {
        propertiesMatrix[ri][itr] = { ...propertiesMatrix[ri][itr], 'w:gridSpan': cspan };
        itr += cspan;
      } else {
        itr++;
      }

    }
  }

  //create a Table Object
  let table = document.createElement('table');
  for (let row of propertiesMatrix) {
    table.insertRow();
    for (let cell of row) {
      delete cell.endAt;
      let newCell = table.rows[table.rows.length - 1].insertCell();
      newCell.innerHTML = `<pre>${JSON.stringify(cell, null, 2)}</pre>`;
    }
  }
  //append the compiled table to the DOM
  document.body.appendChild(table);
})();