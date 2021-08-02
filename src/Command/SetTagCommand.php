<?php

namespace Acme\Command;

use DateTime;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use RetailCrm\Api\Enum\ByIdentifier;
use RetailCrm\Api\Factory\SimpleClientFactory;
use RetailCrm\Api\Model\Entity\Customers\Customer;
use RetailCrm\Api\Model\Entity\Customers\CustomerTag;
use RetailCrm\Api\Model\Request\Customers\CustomersEditRequest;
use RetailCrm\Api\Model\Request\Customers\CustomersRequest;
use Symfony\Component\Console\Command\Command;
use Symfony\Component\Console\Input\InputArgument;
use Symfony\Component\Console\Input\InputInterface;
use Symfony\Component\Console\Output\OutputInterface;
use Throwable;

class SetTagCommand extends Command
{
    const TAG = 'mayorista';

    private $client;

    private $customersAlreadyHaveTag;

    private $notFoundCustomers;

    public function __construct(string $name = null)
    {
        parent::__construct($name);
        $this->customersAlreadyHaveTag = [];
        $this->notFoundCustomers = [];
        $this->client = SimpleClientFactory::createClient();
    }

    protected function configure()
    {
        $this
            ->setName('rcrm:settag')
            ->addArgument('tag', InputArgument::REQUIRED)
        ;
    }

    protected function execute(InputInterface $input, OutputInterface $output)
    {
        $spreadsheet = IOFactory::load("base datos clientes.xlsx");
        $sheet = $spreadsheet->getSheet(0);
        $lastRow = $sheet->getHighestRow();
        $lastColumn = $sheet->getHighestColumn();
        $count = 0;

        for ($row = 2; $row <= $lastRow; $row++) {
//            if ($count > 10) die;
            $rowData = $sheet->rangeToArray('A' . $row . ':' . $lastColumn . $row,NULL,TRUE,FALSE);
            $rowData = array_shift($rowData);
//            dump($rowData);die;
            $cedula = $rowData[0];
            $email = $rowData[6];
            if (!$cedula) {
                $this->notFoundCustomers[] = [
                    'cedula' => $cedula,
                    'email'  => $email
                ];
                continue;
            }
            if (is_string($cedula)) {
                $cedula = trim(str_replace(' ', '', $cedula));
            }
            $cedula = (string)$cedula;

            $customer = $this->findCustomerByCedula($cedula);
//            dump($customer);die;
            if (!$customer && $email) {
                $customer = $this->findCustomerByEmail($email);

            }

            if ($customer) {
                $this->addTagForCustomer($customer);
                $count++;
            } else {
                $this->notFoundCustomers[] = [
                    'cedula' => $cedula,
                    'email'  => $email
                ];
            }
        }

        $this->saveNotFoundToExcel();

        return 0;
    }

    private function saveNotFoundToExcel()
    {
        unlink('notFound.xlsx');
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        foreach ($this->notFoundCustomers as $key => $customer) {
            $sheet->setCellValue('A' . $key, $customer['cedula']);
            $sheet->setCellValue('B' . $key, $customer['email']);
        }

        $writer = new Xlsx($spreadsheet);
        $writer->save('notFound.xlsx');
    }

    private function findCustomerByCedula(string $cedula)
    {
        $request = new CustomersRequest();
        $request->filter->customFields['cedula'] = $cedula;
        $response = $this->client->customers->list($request);
//        dump($response);die;
        $customers = $response->customers;

        return array_shift($customers);
    }

    private function findCustomerByEmail(string $email)
    {
        $request = new CustomersRequest();
        $request->filter->email = $email;
        $response = $this->client->customers->list($request);
        $customers = $response->customers;
        return array_shift($customers);
    }

    private function addTagForCustomer(Customer $customer)
    {
        $customerEditRequest           = new CustomersEditRequest();
        $customerEditRequest->by       = ByIdentifier::ID;
        $customerEditRequest->site     = 'shop-ru';

        $check = $this->checkTag($customer->tags);

        if ($check) {
            $this->customersAlreadyHaveTag[] = $customer;
            return;
        }

        $tag = new CustomerTag();
        $tag->name = self::TAG;
        $customer->tags = array_merge($customer->tags, [$tag]);
        $customer->birthday = null;
        $customerEditRequest->customer = $customer;


        try {
            $response = $this->client->customers->edit($customer->id, $customerEditRequest);

        } catch (Throwable $e){
            dd($e);
        }
    }

    private function checkTag(array $tags): bool
    {
        foreach ($tags as $existingTag) {
            if ($existingTag->name == self::TAG) {
                return true;
            }
        }
        return false;
    }
}