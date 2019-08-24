//*****************************************************************************
// Message board code.
//*****************************************************************************

// Create the TinyMCE control to be used for editing message board posts.
tinyMCE.init({
	apply_source_formatting : true,
	cleanup_on_startup : true,
	content_css : "styles/common.css",
	convert_fonts_to_spans : true,
	doctype : '<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">',
	editor_selector : "mceEditor",
	fix_list_elements : true,
	font_size_style_values : "xx-small,x-small,small,medium,large,x-large,xx-large",
	gecko_spellcheck : true,
	hide_selects_on_submit : true,
	inline_styles : true,
	mode: "textareas",
	plugins : "iespell",
	strict_loading_mode : true,
	theme : "pool",
	theme_pool_buttons1 : "fontselect,fontsizeselect,|,forecolor,backcolor,|,bold,italic,underline,strikethrough,|,link,unlink,image,charmap",
	theme_pool_buttons2 : "bullist,numlist,|,indent,outdent,|,justifyleft,justifyright,justifycenter,justifyfull,|,undo,redo,|,removeformat,cleanup,iespell,code,|,help",
	theme_pool_buttons3 : ""
});
