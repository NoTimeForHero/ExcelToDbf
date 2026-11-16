// Append to a COPY of a company's config.js. Replace markers and coordinates.
app.forms.push({
    name: 'New document type',
    settings: { startY: -1, endX: 6 },
    rules: function() {
        if (!assert(cell(2, 2), 'REPLACE: document identifier')) return;
        const header = findRange(4, 100, 2, 2, 'REPLACE: ID header');
        if (!header) { assert('missing header', 'expected header'); return; }
        context.startY = header.Y + 1;
        assert(cell(header.Y, 5), 'REPLACE: amount header');
    },
    dbfFields: [
        { name: 'ID', length: '8' },
        { name: 'SUMMA', type: 'number', length: '10,2' },
        { name: 'DATEOPL', type: 'date' }
    ],
    beforeWrite: function() {
        context.sum = 0;
        context.totalSeen = false;
    },
    write: function(line) {
        if (line[2] === 'REPLACE: TOTAL marker') {
            context.totalSeen = true;
            context.documentSum = Number(line[5]);
            return stop();
        }
        if (!line[2]) return null;
        const amount = Number(line[5]);
        if (!isFinite(amount) || line[5] === null || line[5] === '') error('Invalid payment amount');
        context.sum += amount;
        return { ID: line[2], SUMMA: amount, DATEOPL: line[6] };
    },
    afterWrite: function() {
        if (!context.totalSeen) error('TOTAL row was not found');
        if (!isFinite(context.documentSum)) error('Invalid TOTAL amount');
        if (Math.abs(context.sum - context.documentSum) > 0.01) error('Payment total mismatch');
    }
});
