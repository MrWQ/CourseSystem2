/*
 * jQuery Plug-In: Table fixed head
 *
 *
 * The MIT License
 *
 * Copyright 2016 Michael Weber <me@michiweber.de>
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 */

(function($){
	$.fn.tfh = function(){

		var method = (arguments.length === 2) ? arguments[0] : ((arguments.length === 1 && typeof arguments[0] === 'string' ? arguments[0] : undefined));
		var options = $.extend({
			trigger: 0,
			top: 0
		},(arguments.length === 2) ? arguments[1] : ((arguments.length === 1 && typeof arguments[0] === 'object' ? arguments[0] : {} )));

		this.width = function(){
			return this.find('thead').attr('data-tmp-width',parseInt(this.find('thead').css('width'))).find('*').each(function(){
				$(this).attr('data-tmp-width',parseInt($(this).css('width')));
			}).end().end();
		};

		this.fix = function(){
			return this.find('.table-fixed-head-thead').css({
				'top': options.top + 'px',
				'position': 'fixed'
			}).end();
		};

		this.clone = function(){
			return this.find('thead').clone(true).prependTo(this).addClass('table-fixed-head-thead').end().end().removeAttr('data-tmp-width').find('*').removeAttr('data-tmp-width').end().end();
		};

		this.build = function(){
			return this.tfh('width').tfh('clone').find('[data-tmp-width]').each(function(){
				$(this).css({
					'width': $(this).data('tmp-width') + 'px',
					'minWidth': $(this).data('tmp-width') + 'px',
					'maxWidth': $(this).data('tmp-width') + 'px',
				});
			}).removeAttr('data-tmp-width').end().tfh('fix', options);
		};

		this.kill = function(){
			this.find('.table-fixed-head-thead').remove();
		};

		this.show = function(){
			return this.addClass('fixed').find('thead').css('visibility','visible').not('.table-fixed-head-thead').css('visibility','hidden').end().end();
		};

		this.hide = function(){
			return this.removeClass('fixed').find('thead').css('visibility','hidden').not('.table-fixed-head-thead').css('visibility','visible').end().end();
		};

		if(method !== undefined){
			return this[method].call($(this));
		} else {
			var table = this.build.call($(this),options);
			if($(document).scrollTop() > options.trigger) {
				table.tfh('show');
			} else {
				table.tfh('hide');
			}
			var resizeTimer;
			var tableScrollLeft = table.position().left;
			$(window).scroll(function(){
				if($(document).scrollTop() > options.trigger) {
					table.tfh('show');
					table.find('.table-fixed-head-thead').css('left',(tableScrollLeft - $(document).scrollLeft()) + 'px');
				} else {
					table.tfh('hide');
				}
			}).resize(function(){
				table.tfh('kill');
				clearTimeout(resizeTimer);
				resizeTimer = setTimeout(function(){
					table.tfh(options);
				}, 250);
			});
		}
	}
	$(document).ready(function(){
		$('table.table-fixed-head').each(function(){
			$(this).tfh({
				trigger: ($(this).data('table-fixed-head-trigger') !== undefined ? $(this).data('table-fixed-head-trigger') : 0),
				top: ($(this).data('table-fixed-head-top') !== undefined ? $(this).data('table-fixed-head-top') : $(this).position().top)
			});
		});
	});
}(jQuery));
