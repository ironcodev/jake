"use strict";

$(function () {
	const hash = $('.page-hash').text().split(',').map(ch => String.fromCharCode(ch)).join('');
	
	$('form').each((i, frm) => {
		const btnSubmit = $(frm).find('.btn-submit');

		if (btnSubmit) {
			$('<input/>')
				.attr('type', 'hidden')
				.attr('name', 'fp')
				.val(hash)
				.appendTo($(frm));

			btnSubmit.on('mouseover', e => {
				btnSubmit.data('ok', '1')
			}).on('mouseout', e => {
				btnSubmit.data('ok', '0')
			});
			
			$(frm).on('submit', e => {
				const ok = btnSubmit.data('ok');

				if (ok != '1') {
					e.preventDefault();
				}
			})
		}
	})
})
