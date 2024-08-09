class UserMailer < ApplicationMailer

	def welcome_mailer
		mail(from: 'nitin.gaurav@reddoorz.com', to: 'nitin.gaurav@reddoorz.com' , subject: 'Welcome Reddoorz')
	end
end
