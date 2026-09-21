import { useContext, useEffect, useState } from 'react';
import AUButton from '@/components/ui/AUButton';
import AUInput from '@/components/ui/AUInput';
import { NextPage } from 'next';
import { clearSession, loginAPI, verifyO365Token } from '@/api/login';
import { setCookie } from 'cookies-next';
import { useRouter } from 'next/router';
import { useForm } from 'react-hook-form';
import { toast } from 'react-toastify';
// import 'react-toastify/dist/ReactToastify.css';
import { useMsal } from '@azure/msal-react';
import { AppContext } from '@/utils/state/app';
import AUModal from '@/components/ui/AUModal';
import { forgotAPI } from '@/api/forgotPassword';
import Image from 'next/image';
import { env } from 'process';
import { EyeIcon } from '@heroicons/react/20/solid';
import { EyeSlashIcon } from '@heroicons/react/20/solid';
interface Props {
}

interface SignInForm {
	username: string
	password: string
}

interface ForgotPasswordForm {
	email: string
}

const Login: NextPage<Props> = (props) => {
	const router = useRouter();
	const { register, handleSubmit, formState: { errors }, setValue, getValues } = useForm<SignInForm>();
	const { register: registerForgotPassword, handleSubmit: handleSubmitForgotPassword, formState: { errors: forgotPasswordErrors }, getValues: getForgotPasswordValues } = useForm<ForgotPasswordForm>();
	const { state: AppState, dispatch: AppDispatch } = useContext(AppContext);
	const [isForgotPasswordModalOpen, setIsForgotPasswordModalOpen] = useState(false);
	const [isForgotPasswordConfirmModalOpen, setIsForgotPasswordConfirmModalOpen] = useState(false);
	const { instance } = useMsal();
	// defaultValues: { username: props.username }
	const [isVisible, setIsVisible] = useState(false);
	const toggleVisibility = () => setIsVisible(!isVisible);

	const onSubmit = async (data: SignInForm) => {
		try {
			const apiResponse = await loginAPI(data);
			// commenting for testing in the QA env and wrote the below code
			if (apiResponse.success) {
            AppDispatch({ type: "SET_USER", payload: apiResponse.data.user });

            // Set user cookie
            const setUserCookiePromise = AppDispatch({ type: "SET_USER_COOKIE", payload: apiResponse.data.user });

            // Set token cookie
            const setTokenCookiePromise = AppDispatch({ type: "SET_TOKEN_COOKIE", payload: apiResponse.data.token });

            // Set login status
            AppDispatch({ type: "SET_LOGIN_TRUE" });

            // Wait for all asynchronous operations to complete before redirecting
            await Promise.all([setUserCookiePromise, setTokenCookiePromise, router.push('/')]);
			toast.success("Logged in successfully")
        }
			else {
				toast.error(apiResponse.message);
			}
		} catch (error) {
			toast.error(`User's name or Password is incorrect`);
		}
	};

	const loginWithO365 = async () => {
		try {
			const loginResponse = await instance.loginPopup();
			const response = await verifyO365Token(loginResponse);
			if (response.success === true) {
				AppDispatch({ type: "SET_USER", payload: response.data.user });

            // Set user cookie
            const setUserCookiePromise = AppDispatch({ type: "SET_USER_COOKIE", payload: response.data.user });

            // Set token cookie
            const setTokenCookiePromise = AppDispatch({ type: "SET_TOKEN_COOKIE", payload: response.data.token });

            // Set login status
            AppDispatch({ type: "SET_LOGIN_TRUE" });

            // Wait for all asynchronous operations to complete before redirecting
            await Promise.all([setUserCookiePromise, setTokenCookiePromise, router.push('/')]);
			toast.success("Logged in successfully")
			}
			else {
				toast.error(response.message);

			}
		} catch (error) {
			// Clear the session keys
			const sessionKeys = ['msal.interaction.status', 'server-telemetry-undefined', `server-telemetry-${env.BACKEND_URL}`]
			await clearSession(sessionKeys);
		}

	};

	const handleForgotPasswordClick = () => {
		setIsForgotPasswordModalOpen(true);
	};

	const handleForgotPasswordClose = () => {
		setIsForgotPasswordModalOpen(false);
	};

	const handleForgotPasswordConfirmClose = () => {
		setIsForgotPasswordConfirmModalOpen(false);
	};

	const handleSubmitForgot = async (data: ForgotPasswordForm) => {
		try {
			const response = await forgotAPI(data)
			if (response.success === true) {
				setIsForgotPasswordModalOpen(false);
				setIsForgotPasswordConfirmModalOpen(true);
			} else {
				toast.error(response.message);
			}
		}
		catch (error) {
			toast.error('An unexpected error occurred');
		}
	}

	return (
		<>
			{/* <ToastContainer /> */}
			<div className='w-full flex justify-center'>
				<div className='mb-2 w-full max-w-[380px]'>
					<div className='flex justify-end mb-2'>
						<Image src="./images/fox-svgrepo-com.svg" alt="Logo" className='h-14' width={100} height={100} />
					</div>
					<h1 className='text-3xl mb-5 text-app-003 font-bold'>Sign In</h1>
					<form onSubmit={handleSubmit((e) => onSubmit(e))}>
						<div className='mb-5 w-full'>
							<AUInput
								// className='w-96'
								placeholder='Email'
								kind='secondary'
								register={register('username', { 
									required: 'Email is required',
									pattern: {
										value: /^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}$/i,
										message: 'Invalid email address'
									}
								})}
								defaultValue={getValues('username')}
								iconType="MAIL"
							/>
							{errors.username && <div className="text-xs text-red-500">{errors.username.message}</div>}
						</div>
						<div className='mb-5 w-full'>
							<AUInput
								// className='w-96'
								placeholder='Password'
								iconType="LOCK"
								type={isVisible ? "text" : "password"}
								kind='secondary'
								register={register('password', { 
									required: 'Password is required',
									minLength: {
										value: 8,
										message: 'Password should be minimum 8 characters long'
									},
								})}
								endContent={<button className="focus:outline-none" type="button" onClick={toggleVisibility}>
								{isVisible ? (
								  <EyeIcon className="text-2xl text-default-400 pointer-events-none" />
								) : (
								  <EyeSlashIcon className="text-2xl text-default-400 pointer-events-none" />
								)}
							  </button>}

							/>
							{errors.password && <div className="text-xs text-red-500">{errors.password.message}</div>}
						</div>
						<div className='flex flex-col md:flex-row mt-5'>
							<div className='mb-3 md:mb-0 md:mr-3'>
								<AUButton type='submit' buttontype='primary' afterImage='RIGHT'>SIGN IN</AUButton>
							</div>
							<div className='md:ml-auto'>
								<AUButton buttontype='outline' onClick={loginWithO365}>SIGN IN WITH O365</AUButton>
							</div>
						</div>
					</form>

					<p className='flex justify-end text-xs mt-2 cursor-pointer text-gray-100' onClick={handleForgotPasswordClick}>Forgot Password</p>
					{/* Forgot Password Modal pop up */}
					<div>
						<AUModal open={isForgotPasswordModalOpen} onClose={handleForgotPasswordClose} title="Forgot Password" size={'lg'}>
							<p>Enter your App' email address below to receive a link to reset your password</p>
							<form onSubmit={handleSubmitForgotPassword(e => handleSubmitForgot(e))}>
								<div className="mb-4">
									<AUInput
										// className='w-96'
										placeholder='Email'
										kind='primary'
										register={registerForgotPassword('email', { 
											required: 'Email is required',
											pattern: {
												value: /^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}$/i,
												message: 'Invalid email address'
											}
										})}
										defaultValue={getForgotPasswordValues('email')}
										iconType="MAIL"
									/>
									{forgotPasswordErrors.email && <div className="text-xs text-red-500">{forgotPasswordErrors.email.message}</div>}
								</div>
								<div className='flex justify-end  mt-4 mb-2'>
									<div className='mr-2'>
										<AUButton buttontype='danger' color="danger" onClick={handleForgotPasswordClose}>Close</AUButton>
									</div>
									<div>
										<AUButton type='submit' buttontype='default' afterImage='RIGHT' className='w-5' >Request Password Reset Email</AUButton>
									</div>
								</div>
							</form>
						</AUModal>
					</div>

					{/* Forgot Password confirmation Modal pop up */}
					<div>
						<AUModal open={isForgotPasswordConfirmModalOpen} onClose={handleForgotPasswordConfirmClose} title="Info" size={'sm'}>
							<form onSubmit={handleSubmitForgotPassword(e => handleSubmitForgot(e))}>
								<div className="mb-4">
									<p>The one time password reset link has been sent to your email account.</p>
								</div>
								<div className='flex justify-end  mt-4 mb-2'>
									<div className='mr-2'>
										<AUButton buttontype='danger' color="danger" onClick={handleForgotPasswordConfirmClose}>Close</AUButton>
									</div>
								</div>
							</form>
						</AUModal>
					</div>
				</div>
			</div>
		</>
	);
};

export default Login;