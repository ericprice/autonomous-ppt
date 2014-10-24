tell application "Microsoft PowerPoint"
	activate
	# add create new presentation here eventually
	tell active presentation
		make new slide at end with properties {layout:slide layout text slide}
	end tell
	set settings to slide show settings of active presentation
	run slide show settings
	
	set textObjects to {"SYNERGY", "DELIVERABLES", "ANALYTICS", "BALLPARK FIGURE", "BRANDS", "CAPABILITIES", "DOWNSIZING", "UPSIZING", "CORE COMPETENCIES", "BRICK-AND-MORTAR", "EYEBALLS", "GRANULARITY", "LONG TAIL", "LOW HANGING FRUIT", "MINDSHARE", "OFFSHORING", "ROI", "CONVERGENCE", "CLOUD", "WEBINARS", "REAL-TIME", "SCALABILITY", "STRUTS", "PROCESS", "DEVELOPMENT", "DONORS", "TRANSMEDIA", "MASHUPS", "INNOVATION", "ALIGNMENT", "DIVERSITY", "EMPOWERMENT", "EXIT STRATEGIES", "GENERATION X", "AGENDAS", "MEETINGS", "GENERATION WHY", "PARADIGMS", "WIN-WIN", "SLIDESHOWS", "ORACLES", "AUTONOMY", "TRANSITIONS", "CHARTS", "THEMES", "SHAPES", "COLORS", "POWERPOINT", "POWERPOINT", "POWERPOINT", "POWERPOINT", "DOWNSIZING", "UPSIZING", "CROWDSOURCING", "NOMADS", "DATA", "SKILLS", "PLATFORMS", "CURATION", "CAPITAL", "NICHES", "LOCALIZATION", "IDEATION", "IMPLEMENTATION", "USERS", "DESIGN", "VISIONARIES", "TEAMS", "SKILLS", "COMPUTERS", "$$$", "$$$", "ANIMALS", "CLIENTS", "CREATIVES"}
	
	set textVerbs to {"OBSERVE", "SEE", "HEAR", "FEEL", "PRESENT", "WISH FOR", "USE", "SHOW", "BROADCAST", "ILLUMINATE", "BRING SHAPE TO", "RESPOND TO", "ENABLE"}
	
	set textAdjectives to {"WARM", "COLD", "LOUD", "QUIET", "SEXY", "STRONG", "ESSENTIAL", "BORING", "DULL", "EXCITING", "FRESH", "NEW", "OPPORTUNE", "RELATIVE", "HOLISTIC", "COMPUTING", "GENERATING", "FINAL", "FLUENT", "FIRST", "MID-LEVEL", "AUTONOMOUS", "ENRICHING", "SINGULAR", "BRANDED", "WONDROUS", "ENRICHING", "COST-EFFECTIVE", "PRIMAL", "TWISTED", "WORLD-CLASS", "FRANK", "ESCALATING", "ANONYMOUS", "FREQUENT", "INFINITE", "MISSION-CRITICAL", "HIDDEN", "OBVIOUS", "TOP-DOWN", "BOTTOM-UP", "LONG", "PLAIN", "POWERFUL", "VAST", "MYSTERIOUS", "AGREEABLE", "SHALLOW", "STEEP", "WIDE", "HIGH", "HIGH", "LOW", "MASSIVE", "MASSIVE", "GIANT", "MODERN", "MODERN", "SWIFT", "RAPID", "RAPID", "BITTER", "WEAK", "HOT", "HOT", "FULL", "FULL", "SUBSTANTIAL", "DRAB", "CLEAR", "CLEAN", "UNUSUAL", "REAL", "REAL", "WILD", "FAIR", "RETRO", "STALE", "COOL", "ABUNDANT", "ABUNDANT", "STRATEGIC", "STRATEGIC", "FREEMIUM", "FREEMIUM", "SOCIAL", "REPURPOSED", "UBIQUITOUS", "UBIQUITOUS", "DYNAMIC", "RANDOM"}
	
	set textSelf to {"I", "WE", "YOU"}
	
	set textConnectors to {"AS", "IN", "WITH", "AND", "IS", "WILL BE"}
	
	set bgTextures to {texture blue tissue paper, texture bouquet, texture brown marble, texture canvas, texture cork, texture denim, texture fish fossil, texture granite, texture green marble, texture medium wood, texture newsprint, texture oak, texture paper bag, texture papyrus, texture parchment, texture pink tissue paper, texture purple mesh, texture recycled paper, texture sand, texture stationery, texture walnut, texture water droplets, texture white marble, texture woven mat}
	
	set entryEffects to {entry effect checkerboard across, entry effect checkerboard down, entry effect blinds vertical, entry effect dissolve, entry effect comb horizontal, entry effect comb vertical, entry effect box in, entry effect box out}
	
	set miscShapes to {autoshape cross, autoshape can, autoshape bevel, autoshape smiley face, autoshape heart, autoshape sun, autoshape right arrow, autoshape up arrow, autoshape left right arrow, autoshape quad arrow, autoshape bent arrow, autoshape five point star}
	
	set arrowShapes to {autoshape right arrow, autoshape up arrow, autoshape left right arrow, autoshape quad arrow, autoshape bent arrow, autoshape left up arrow, autoshape curved right arrow, autoshape curved up arrow, autoshape striped right arrow, autoshape left arrow, autoshape down arrow, autoshape up down arrow, autoshape left right up arrow, autoshape curved left arrow, autoshape curved down arrow, autoshape circular arrow}
	
	set pptPicFolder to "Macintosh HD:Users:ericprice:Desktop:art:"
	tell application "Finder"
		set pptPics to name of every file of folder "art" of folder "Desktop" of home
	end tell
	set pptPicsCount to count of pptPics
	
	set slideCount to 1
	
	repeat
		
		# Create new slide
		tell active presentation
			make new slide at end with properties {layout:slide layout text slide}
			set entry effect of slide show transition of slide slideCount to some item of entryEffects
		end tell
		
		# Grab pictures
		set picFolder to "Macintosh HD:Users:ericprice:Desktop:pics:"
		tell application "Finder"
			set pics to name of every file of folder "pics" of folder "Desktop" of home
		end tell
		set picCount to count of pics
		
		set pathForward to random number from 1 to 3
		
		tell slide slideCount of active presentation
			
			# CORPORATE SPEAK
			if pathForward is 1 then
				
				# Add background
				set follow master background to false
				preset textured background texture some item of bgTextures
				set bgMask to make new shape at beginning with properties {auto shape type:autoshape rectangle, left position:20, top:20, width:680, height:500}
				set fore color of fill format of bgMask to ({255, 255, 255} as RGB color)
				set back color of fill format of bgMask to ({255, 255, 255} as RGB color)
				set fore color of line format of bgMask to ({255, 255, 255} as RGB color)
				set back color of line format of bgMask to ({255, 255, 255} as RGB color)
				set visible of shadow format of bgMask to false
				
				set pathForwardPics to random number from 1 to 2
				
				if pathForwardPics is 1 then
					# Add a picture, randomly position (based on 320x240 image in 720x540 canvas)
					set picXRandom to random number from 30 to 400
					set picYRandom to random number from 30 to 300
					set pic to picFolder & some item of pics
					set displayedPic to make new picture at end with properties {top:picYRandom, left position:picXRandom, lock aspect ratio:true, file name:pic}
					tell displayedPic
						scale height factor 0.25 scale scale from top left with relative to original size
						scale width factor 0.25 scale scale from top left with relative to original size
					end tell
				end if
				if pathForwardPics is 2 then
					# Add a PPT-art picture
					set picXRandom to random number from 30 to 300
					set picYRandom to random number from 30 to 200
					set pic to pptPicFolder & some item of pptPics
					set displayedPic to make new picture at end with properties {top:picYRandom, left position:picXRandom, lock aspect ratio:true, file name:pic}
					tell displayedPic
						scale height factor 0.4 scale scale from top left with relative to original size
						scale width factor 0.4 scale scale from top left with relative to original size
					end tell
				end if
				
				set shapeXRandom to random number from 30 to 660
				set shapeYRandom to random number from 30 to 480
				set foreMiscShape1 to make new shape at end with properties {auto shape type:some item of miscShapes, left position:shapeXRandom, top:shapeYRandom, width:random number from 20 to 60, height:random number from 10 to 100}
				set fore color of fill format of foreMiscShape1 to {random number from 0 to 255, random number from 0 to 255, random number from 0 to 255}
				set back color of fill format of foreMiscShape1 to {random number from 0 to 255, random number from 0 to 255, random number from 0 to 255}
				set fore color of line format of foreMiscShape1 to {random number from 0 to 255, random number from 0 to 255, random number from 0 to 255}
				set back color of line format of foreMiscShape1 to {random number from 0 to 255, random number from 0 to 255, random number from 0 to 255}
				set rotation of foreMiscShape1 to random number from -30 to 30
				set visible of shadow format of foreMiscShape1 to false
				
				# Add text
				set pathForwardMeta to random number from 1 to 2
				set shapeXRandom to random number from 30 to 550
				set shapeYRandom to random number from 30 to 480
				set meta to make new text box at end with properties {text orientation:horizontal, left position:shapeXRandom, top:shapeYRandom, width:300, height:100}
				if pathForwardMeta is 1 then
					set content of text range of text frame of meta to "75¡"
				end if
				if pathForwardMeta is 2 then
					set theHour to get the (hours of (current date)) as string
					set theMinute to get the (minutes of (current date)) as string
					set minuteLength to (length of theMinute)
					if minuteLength < 2 then
						set theMinute to "0" & theMinute
					end if
					set content of text range of text frame of meta to theHour & ":" & theMinute
				end if
				
				tell font of text range of text frame of meta
					set font name to "Franklin Gothic Std ExtraCond"
					set font size to 40
					set font color to {0, 0, 0}
				end tell
				
				set headline to some item of textAdjectives & " " & some item of textObjects
				set content of text range of text frame of shape 1 to headline
				tell font of text range of text frame of shape 1
					set font name to "Times New Roman MT Std Cond"
					set font size to 85
					set font color to {random number from 0 to 255, random number from 0 to 255, random number from 0 to 255}
				end tell
				set top of shape 1 to 0
				set height of shape 1 to 540
				z order shape 1 z order position bring shape to front
				
				# Add a shape, randomly position (based on 60x60 image in 720x540 canvas)
				set shapeXRandom to random number from 30 to 660
				set shapeYRandom to random number from 30 to 480
				set arrow1 to make new shape at end with properties {auto shape type:some item of arrowShapes, left position:shapeXRandom, top:shapeYRandom, width:random number from 20 to 60, height:random number from 10 to 100}
				set fore color of fill format of arrow1 to {random number from 0 to 255, random number from 0 to 255, random number from 0 to 255}
				set back color of fill format of arrow1 to {random number from 0 to 255, random number from 0 to 255, random number from 0 to 255}
				set fore color of line format of arrow1 to {random number from 0 to 255, random number from 0 to 255, random number from 0 to 255}
				set back color of line format of arrow1 to {random number from 0 to 255, random number from 0 to 255, random number from 0 to 255}
				set rotation of arrow1 to random number from -30 to 30
				set visible of shadow format of arrow1 to false
				
				set shapeXRandom to random number from 30 to 660
				set shapeYRandom to random number from 30 to 480
				set arrow2 to make new shape at end with properties {auto shape type:some item of arrowShapes, left position:shapeXRandom, top:shapeYRandom, width:random number from 20 to 60, height:random number from 10 to 100}
				set fore color of fill format of arrow2 to {random number from 0 to 255, random number from 0 to 255, random number from 0 to 255}
				set back color of fill format of arrow2 to {random number from 0 to 255, random number from 0 to 255, random number from 0 to 255}
				set fore color of line format of arrow2 to {random number from 0 to 255, random number from 0 to 255, random number from 0 to 255}
				set back color of line format of arrow2 to {random number from 0 to 255, random number from 0 to 255, random number from 0 to 255}
				set rotation of arrow2 to random number from -30 to 30
				set visible of shadow format of arrow2 to false
				
			end if
			
			# CORPORATE SPEAK LIST
			if pathForward is 2 then
				
				# Add background
				set follow master background to false
				preset textured background texture some item of bgTextures
				set bgMask to make new shape at beginning with properties {auto shape type:autoshape rectangle, left position:20, top:20, width:680, height:500}
				set fore color of fill format of bgMask to ({255, 255, 255} as RGB color)
				set back color of fill format of bgMask to ({255, 255, 255} as RGB color)
				set fore color of line format of bgMask to ({255, 255, 255} as RGB color)
				set back color of line format of bgMask to ({255, 255, 255} as RGB color)
				set visible of shadow format of bgMask to false
				
				set shapeXRandom to random number from 30 to 660
				set shapeYRandom to random number from 30 to 480
				set foreMiscShape1 to make new shape at end with properties {auto shape type:some item of miscShapes, left position:shapeXRandom, top:shapeYRandom, width:random number from 20 to 60, height:random number from 10 to 100}
				set fore color of fill format of foreMiscShape1 to {random number from 0 to 255, random number from 0 to 255, random number from 0 to 255}
				set back color of fill format of foreMiscShape1 to {random number from 0 to 255, random number from 0 to 255, random number from 0 to 255}
				set fore color of line format of foreMiscShape1 to {random number from 0 to 255, random number from 0 to 255, random number from 0 to 255}
				set back color of line format of foreMiscShape1 to {random number from 0 to 255, random number from 0 to 255, random number from 0 to 255}
				set rotation of foreMiscShape1 to random number from -30 to 30
				set visible of shadow format of foreMiscShape1 to false
				
				set headline to some item of textObjects & "
" & some item of textObjects & " " & some item of textConnectors & " " & some item of textObjects & "
" & some item of textSelf & " " & some item of textVerbs & " " & some item of textObjects
				set content of text range of text frame of shape 2 to headline
				tell font of text range of text frame of shape 2
					set font name to "Times New Roman MT Std Cond"
					set font size to 50
					set font color to {random number from 0 to 255, random number from 0 to 255, random number from 0 to 255}
				end tell
				set top of shape 2 to 40
				set height of shape 2 to 500
				z order shape 2 z order position bring shape to front
				
				set shapeXRandom to random number from 30 to 660
				set shapeYRandom to random number from 30 to 480
				set miscShape1 to make new shape at end with properties {auto shape type:some item of miscShapes, left position:shapeXRandom, top:shapeYRandom, width:random number from 20 to 60, height:random number from 10 to 100}
				set fore color of fill format of miscShape1 to {random number from 0 to 255, random number from 0 to 255, random number from 0 to 255}
				set back color of fill format of miscShape1 to {random number from 0 to 255, random number from 0 to 255, random number from 0 to 255}
				set fore color of line format of miscShape1 to {random number from 0 to 255, random number from 0 to 255, random number from 0 to 255}
				set back color of line format of miscShape1 to {random number from 0 to 255, random number from 0 to 255, random number from 0 to 255}
				set rotation of miscShape1 to random number from -30 to 30
				set visible of shadow format of miscShape1 to false
				
				set shapeXRandom to random number from 30 to 660
				set shapeYRandom to random number from 30 to 480
				set miscShape2 to make new shape at end with properties {auto shape type:some item of miscShapes, left position:shapeXRandom, top:shapeYRandom, width:random number from 20 to 60, height:random number from 10 to 100}
				set fore color of fill format of miscShape2 to {random number from 0 to 255, random number from 0 to 255, random number from 0 to 255}
				set back color of fill format of miscShape2 to {random number from 0 to 255, random number from 0 to 255, random number from 0 to 255}
				set fore color of line format of miscShape2 to {random number from 0 to 255, random number from 0 to 255, random number from 0 to 255}
				set back color of line format of miscShape2 to {random number from 0 to 255, random number from 0 to 255, random number from 0 to 255}
				set rotation of miscShape2 to random number from -30 to 30
				set visible of shadow format of miscShape2 to false
				
			end if
			
			# CORPORATE SPEAK 2
			if pathForward is 3 then
				
				# Add background
				set follow master background to false
				set bgMask to make new shape at beginning with properties {auto shape type:autoshape rectangle, left position:0, top:0, width:720, height:540}
				
				set pathForwardColors to random number from 1 to 5
				
				if pathForwardColors is 1 then
					set fore color of fill format of bgMask to ({0, 0, 207} as RGB color)
					set back color of fill format of bgMask to ({0, 0, 207} as RGB color)
					set fore color of line format of bgMask to ({0, 0, 207} as RGB color)
					set back color of line format of bgMask to ({0, 0, 207} as RGB color)
				end if
				
				if pathForwardColors is 2 then
					set fore color of fill format of bgMask to ({214, 0, 0} as RGB color)
					set back color of fill format of bgMask to ({214, 0, 0} as RGB color)
					set fore color of line format of bgMask to ({214, 0, 0} as RGB color)
					set back color of line format of bgMask to ({214, 0, 0} as RGB color)
				end if
				
				if pathForwardColors is 3 then
					set fore color of fill format of bgMask to ({0, 255, 0} as RGB color)
					set back color of fill format of bgMask to ({0, 255, 0} as RGB color)
					set fore color of line format of bgMask to ({0, 255, 0} as RGB color)
					set back color of line format of bgMask to ({0, 255, 0} as RGB color)
				end if
				
				if pathForwardColors is 4 then
					set fore color of fill format of bgMask to ({0, 0, 0} as RGB color)
					set back color of fill format of bgMask to ({0, 0, 0} as RGB color)
					set fore color of line format of bgMask to ({0, 0, 0} as RGB color)
					set back color of line format of bgMask to ({0, 0, 0} as RGB color)
				end if
				
				if pathForwardColors is 5 then
					set fore color of fill format of bgMask to ({0, 0, 207} as RGB color)
					set back color of fill format of bgMask to ({0, 0, 207} as RGB color)
					set fore color of line format of bgMask to ({0, 0, 207} as RGB color)
					set back color of line format of bgMask to ({0, 0, 207} as RGB color)
				end if
				
				set visible of shadow format of bgMask to false
				
				set headline to some item of textAdjectives & " " & some item of textObjects
				set content of text range of text frame of shape 1 to headline
				tell font of text range of text frame of shape 1
					set font name to "Times New Roman MT Std Cond"
					set font size to 90
					if pathForwardColors is 1 then
						set font color to {150, 255, 175}
					end if
					if pathForwardColors is 2 then
						set font color to {0, 0, 0}
					end if
					if pathForwardColors is 3 then
						set font color to {0, 0, 0}
					end if
					if pathForwardColors is 4 then
						set font color to {0, 16, 169}
					end if
					if pathForwardColors is 5 then
						set font color to {255, 0, 0}
					end if
				end tell
				set top of shape 1 to 0
				set height of shape 1 to 540
				z order shape 1 z order position bring shape to front
				
			end if
			
		end tell
		
		# How long each slide is displayed
		delay 1
		
		# Wait for the second slide to load before continuing so we're always one ahead in generation
		if slideCount is greater than 3 then
			go to next slide slide show view of slide show window 1
		end if
		
		set slideCount to slideCount + 1
	end repeat
	
end tell